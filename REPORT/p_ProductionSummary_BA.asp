<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%	
	Dim rsEPS, strEPS
	Dim v_txtyear, v_txtmonth, v_txttodaydate, v_txtDivision, v_txtBlok, v_txtPlantingDate, NO, cClass, v_txtFcba

	v_txtyear			=	trim(request.Form("txtyear"))
	v_txtmonth			=	trim(request.Form("txtmonth"))
	v_txttodaydate		=	trim(request.Form("txttodaydate"))
	v_txtDivision		=	trim(request.Form("txtDivision"))
	v_txtBlok			=	trim(request.Form("txtBlok"))
	v_txtPlantingDate	=	trim(request.Form("txtPlantingDate"))
	v_txtFcba 		    =   trim(request.Form("txtFcba"))
	
	
	strEPS  =  " SELECT * " & _ 
				   " FROM (SELECT   fcba, division, fieldcode, fieldname, plantingyear, " & _
								 " SUM (hi) hi, SUM (shi) shi, SUM (bi) bi, SUM (sbi) sbi, " & _
								 " SUM (tonase_hi) tonase_hi, SUM (tonase_shi) tonase_shi, " & _
								 " SUM (tonase_bi) tonase_bi, SUM (tonase_sbi) tonase_sbi, " & _
								 " SUM (pom_hi) pom_hi, SUM (pom_shi) pom_shi, " & _
								 " SUM (pom_bi) pom_bi, SUM (pom_sbi) pom_sbi, " & _
								 " SUM (pom_tonase_hi) pom_tonase_hi, " & _
								 " SUM (pom_tonase_shi) pom_tonase_shi, " & _
								 " SUM (pom_tonase_bi) pom_tonase_bi, " & _
								 " SUM (pom_tonase_sbi) pom_tonase_sbi, SUM (restan) restan " & _
							" FROM (SELECT fcba, division, fieldcode, fieldname, plantingyear, " & _
										 " totaldailyproduction hi, 0 shi, 0 bi, 0 sbi, " & _
										 " totaldailyprodtonase tonase_hi, 0 tonase_shi, " & _
										 " 0 tonase_bi, 0 tonase_sbi, 0 pom_hi, 0 pom_shi, " & _
										 " 0 pom_bi, 0 pom_sbi, 0 pom_tonase_hi, " & _
										 " 0 pom_tonase_shi, 0 pom_tonase_bi, 0 pom_tonase_sbi, " & _
										 " 0 restan " & _
									" FROM view_dailyharvesting_prod " & _
								   " WHERE fcba = '" & v_txtFcba & "' " & _
									 " AND accountingyear = '" & v_txtyear & "' " & _
									 " AND accountingperiod = '" & v_txtmonth & "' " & _
									 " AND fddate = TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT   fcba, division, fieldcode, fieldname, plantingyear, " & _
										   " 0 hi, SUM (totaldailyproduction) shi, 0 bi, 0 sbi, " & _
										   " 0 tonase_hi, SUM (totaldailyprodtonase) tonase_shi, " & _
										   " 0 tonase_bi, 0 tonase_sbi, 0 pom_hi, 0 pom_shi, " & _
										   " 0 pom_bi, 0 pom_sbi, 0 pom_tonase_hi, " & _
										   " 0 pom_tonase_shi, 0 pom_tonase_bi, " & _
										   " 0 pom_tonase_sbi, 0 restan " & _
									  " FROM view_dailyharvesting_prod " & _
									 " WHERE fcba = '" & v_txtFcba& "' " & _
									   " AND accountingyear = '" & v_txtyear & "' " & _
									   " AND accountingperiod = '" & v_txtmonth & "' " & _
									   " AND fddate <= TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
								  " GROUP BY fcba, " & _
										   " accountingyear, " & _
										   " accountingperiod, " & _
										   " division, " & _
										   " fieldcode, " & _
										   " fieldname, " & _
										   " plantingyear " & _
								  " UNION " & _
								  " SELECT   fcba, division, fieldcode, fieldname, plantingyear, " & _
										   " 0 hi, 0 shi, 0 bi, SUM (totaldailyproduction) sbi, " & _
										   " 0 tonase_hi, 0 tonase_shi, 0 tonase_bi, " & _
										   " SUM (totaldailyprodtonase) tonase_sbi, 0 pom_hi, " & _
										   " 0 pom_shi, 0 pom_bi, 0 pom_sbi, 0 pom_tonase_hi, " & _
										   " 0 pom_tonase_shi, 0 pom_tonase_bi, " & _
										   " 0 pom_tonase_sbi, 0 restan " & _
									  " FROM view_dailyharvesting_prod " & _
									 " WHERE fcba = '" & v_txtFcba& "' " & _
									   " AND accountingyear = '" & v_txtyear & "' " & _
									   " AND fddate <= TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
								  " GROUP BY fcba, division, fieldcode, fieldname, plantingyear " & _
								  " UNION " & _
								  " SELECT fcba, division, fieldcode, fieldname, plantingyear, " & _
										 " 0 hi, 0 shi, 0 bi, 0 sbi, 0 tonase_hi, 0 tonase_shi, " & _
										 " 0 tonase_bi, 0 tonase_sbi, totaldailyreceive pom_hi, " & _
										 " 0 pom_shi, 0 pom_bi, 0 pom_sbi, " & _
										 " totaldailyreceivetonase1 pom_tonase_hi, " & _
										 " 0 pom_tonase_shi, 0 pom_tonase_bi, 0 pom_tonase_sbi, " & _
										 " 0 restan " & _
									" FROM view_dailyharvesting_mill " & _
								   " WHERE fcba = '" & v_txtFcba& "' " & _
									 " AND accountingyear = '" & v_txtyear & "' " & _
									 " AND accountingperiod = '" & v_txtmonth & "' " & _
									 " AND receptiondate = TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
								  " UNION " & _
								  " SELECT   fcba, division, fieldcode, fieldname, plantingyear, " & _
										   " 0 hi, 0 shi, 0 bi, 0 sbi, 0 tonase_hi, " & _
										   " 0 tonase_shi, 0 tonase_bi, 0 tonase_sbi, 0 pom_hi, " & _
										   " SUM (totaldailyreceive) pom_shi, 0 pom_bi, " & _
										   " 0 pom_sbi, 0 pom_tonase_hi, " & _
										   " SUM (totaldailyreceivetonase1) pom_tonase_shi, " & _
										   " 0 pom_tonase_bi, 0 pom_tonase_sbi, 0 restan " & _
									  " FROM view_dailyharvesting_mill " & _
									 " WHERE fcba = '" & v_txtFcba& "' " & _
									   " AND accountingyear = '" & v_txtyear & "' " & _
									   " AND accountingperiod = '" & v_txtmonth & "' " & _
									   " AND receptiondate <=  TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
								  " GROUP BY fcba, " & _
										   " accountingyear, " & _
										   " accountingperiod, " & _
										   " division, " & _
										   " fieldcode, " & _
										   " fieldname, " & _
										   " plantingyear " & _
								  " UNION " & _
								  " SELECT   fcba, division, fieldcode, fieldname, plantingyear, " & _
										   " 0 hi, 0 shi, 0 bi, 0 sbi, 0 tonase_hi, " & _
										   " 0 tonase_shi, 0 tonase_bi, 0 tonase_sbi, 0 pom_hi, " & _
										   " 0 pom_shi, 0 pom_bi, " & _
										   " SUM (totaldailyreceive) pom_sbi, 0 pom_tonase_hi, " & _
										   " 0 pom_tonase_shi, 0 pom_tonase_bi, " & _
										   " SUM (totaldailyreceivetonase1) pom_tonase_sbi, " & _
										   " 0 restan " & _
									  " FROM view_dailyharvesting_mill " & _
									 " WHERE fcba = '" & v_txtFcba & "' " & _
									   " AND accountingyear = '" & v_txtyear & "' " & _
									   " AND receptiondate <= TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
								  " GROUP BY fcba, division, fieldcode, fieldname, plantingyear " & _
								  " UNION " & _
								  " select FCBA, DIVISION, FIELDCODE, FIELDNAME, PLANTINGYEAR,  " & _
										   " 0 hi, 0 shi, 0 bi, 0 sbi, 0 tonase_hi, " & _
										   " 0 tonase_shi, 0 tonase_bi, 0 tonase_sbi, 0 pom_hi, " & _
										   " 0 pom_shi, 0 pom_bi, " & _
										   " 0 pom_sbi, 0 pom_tonase_hi, " & _
										   " 0 pom_tonase_shi, 0 pom_tonase_bi, " & _
										   " 0 pom_tonase_sbi, " & _
										   " sum(variantemp) + sum(restantemp) restan " & _
										" from (	 " & _
											" select FCBA, DIVISION, FIELDCODE, FIELDNAME, PLANTINGYEAR,  varian variantemp,  " & _
											" 0 restantemp from view_dailyharvesting_varian " & _
											" where " & _
											" fcba = '" & v_txtFcba& "' and " & _
											" DATEPROD = TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
											" UNION " & _
											" select FCBA, DIVISION, FIELDCODE, FIELDNAME, PLANTINGYEAR,  0 variantemp,  " & _
											" sum(varian) restantemp from view_dailyharvesting_varian " & _
											" where " & _
											" fcba = '" & Session("fcBA")& "' and " & _
											" DATEPROD < TO_DATE ('" & v_txttodaydate & "', 'MM/DD/YYYY') " & _
											" group by FCBA, DIVISION, FIELDCODE, FIELDNAME, PLANTINGYEAR " & _
										" )  " & _
										" group by FCBA, DIVISION, FIELDCODE, FIELDNAME, PLANTINGYEAR " & _
								  " ) " & _
						" GROUP BY fcba, division, fieldcode, fieldname, plantingyear)  " & _
				 " WHERE fcba = '" & v_txtFcba& "' " 
				if len(trim(v_txtDivision)) > 0 then
					strEPS = strEPS & " AND DIVISION = '" & v_txtDivision & "'"
				end if
				if len(trim(v_txtBlok)) > 0 then
					strEPS = strEPS & " AND FIELDCODE = '" & v_txtBlok & "'"
				end if
				if len(trim(v_txtPlantingDate)) > 0 then
					strEPS = strEPS & " AND PLANTINGYEAR = " & cint(v_txtplantingdate)
				end if




Set rsEPS = Server.CreateObject("ADODB.Recordset")
rsEPS.Open strEPS, connstr, adOpenStatic 


%>

<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPS Reporting - Production Summary</title>
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
	<!--#include file="menutop.asp"-->
	<div id="header-wrapper">
		<div id="header">
			<h4>PT. <%=Session("COMPANY")%></h4>
			<h5><%=Session("BA")%></h5>
		</div>
		<div id="subheader"class="centered">
			<h2>LAPORAN HARIAN PRODUKSI PER KEBUN</h2>
			<% if v_txttodaydate <> "" and v_txtFcba <> "" then %>
				<h3>PERIODE : <%=v_txttodaydate%>, KEBUN <%=v_txtFcba%> </h3>
			<% end if %>
		</div>
	</div>

<%
	if len(trim(request("txtYear"))) > 0 then
		v_txtyear = request.Form("txtYear")
	else
		v_txtyear = year(Now)	 
	end if

	if len(trim(request("txtmonth"))) > 0 then
		v_txtmonth = request.Form("txtmonth")
	else
		v_txtmonth = Month(Now)	 
	end if
%>
<table id="form-table" class="centered-margin">
	<tr>
		<th>TAHUN</th>
		<th>BULAN</th>
		<th>TANGGAL</th>
		<th>KEBUN</th>
		<th>AFDELING</th>
		<th>TAHUN TANAM</th>
		<th>BLOK</th>
		<th>&nbsp;</th>
	</tr>
	<tr>
		<form name="frmEntry" method="post" action="p_ProductionSummary_BA.asp">
		<td align="center"><input type="text" name="txtyear" value="<%=v_txtyear%>"></td>
		<td align="center"><input type="text" name="txtmonth" value="<%=v_txtmonth%>"></td>
		<td align="center">
		<% if v_txttodaydate <> "" and v_txttodaydate = v_txttodaydate then %>
			<input type="text" maxlength="10" name="txttodaydate" size="10" value="<%=v_txttodaydate%>" title="Type Date mm/dd/yyyy" >
		<% else%>
			<input type="text" maxlength="10" name="txttodaydate" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			<a href="javascript:show_calendar('frmEntry.txttodaydate');" onMouseOver="window.status='Date Picker';return true;" onmouseout="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>
		</td>
		
		
		<td align="center">
		<select name="txtFcba">	
		<option value="">--NONE--</option>		
		<%
			Dim rsFcba
			Set rsFcba = Server.CreateObject("ADODB.Recordset")
			rsFcba.CursorLocation = adUseClient
			strsql =  " SELECT fccode FROM businessunit "
			rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
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
		
		
		
		<td align="center">
		<select name="txtDivision">	
				<option value="">--ALL--</option>
		<%
			Dim rsDivision
			Set rsDivision = Server.CreateObject("ADODB.Recordset")
			rsDivision.CursorLocation = adUseClient
			strsql =  " SELECT DISTINCT Division as Division FROM FIELD WHERE FCBA = '" & Session("fcBA") & "'"
			rsDivision.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsDivision.eof then
				DO while not rsDivision.eof
				if v_txtDivision = rsDivision("Division") then
		%>			<option value="<%=rsDivision("Division")%>" Selected><%=rsDivision("Division")%></option>
				<%else%>
					<option value="<%=rsDivision("Division")%>"><%=rsDivision("Division")%></option>
				<%end if%>
		<%			rsDivision.movenext
				loop
			End if
			rsDivision.Close
			set rsDivision = nothing
		%>
		</select>		
		</td>


		<td align="center">
		<select name="txtPlantingDate">	
				<option value="">--ALL--</option>
		<%
			Dim rsPlantingDate, strsql
			Set rsPlantingDate = Server.CreateObject("ADODB.Recordset")
			rsPlantingDate.CursorLocation = adUseClient
			strsql =  " SELECT DISTINCT to_char(PLANTINGDATE, 'yyyy') as PlantingDate FROM FIELD WHERE FCBA = '" & Session("fcBA") & "'"
			rsPlantingDate.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsPlantingDate.eof then
				DO while not rsPlantingDate.eof
				if v_txtPlantingDate = rsPlantingDate("PlantingDate") then
		%>			<option value="<%=rsPlantingDate("PlantingDate")%>" Selected><%=rsPlantingDate("PlantingDate")%></option>
				<%else%>
					<option value="<%=rsPlantingDate("PlantingDate")%>"><%=rsPlantingDate("PlantingDate")%></option>
				<%end if%>
		<%			rsPlantingDate.movenext
				loop
			End if
			rsPlantingDate.Close
			set rsPlantingDate = nothing
		%>
		</select>		
		</td>
		<td align="center">
		<select name="txtBlok">	
				<option value="">--ALL--</option>
		<%
			Dim rsBlok
			Set rsBlok = Server.CreateObject("ADODB.Recordset")
			rsBlok.CursorLocation = adUseClient
			strsql =  " SELECT DISTINCT fccode as Blok FROM FIELD WHERE FCBA = '" & Session("fcBA") & "'"
			rsBlok.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
			if not rsBlok.eof then
				DO while not rsBlok.eof
				if v_txtBlok = rsBlok("Blok") then
		%>			<option value="<%=rsBlok("Blok")%>" Selected><%=rsBlok("Blok")%></option>
				<%else%>
					<option value="<%=rsBlok("Blok")%>"><%=rsBlok("Blok")%></option>
				<%end if%>
		<%			rsBlok.movenext
				loop
			End if
			rsBlok.Close
			set rsBlok = nothing
		%>
		</select>		
		</td>
		<td align="center"><input type="submit" value="View" /></td>
		</form>
	</tr>
</table>

<table border="0" width="100%" id="product-table">
	<tr>
		<th rowspan="4">NO</th>
		<th rowspan="4">KEBUN</th>
		<th rowspan="4">AFDELING</th>
		<th rowspan="4">TT</th>
		<th rowspan="4">BLOK</th>
		<th colspan="8">PRODUKSI TBS KEBUN</th>
		<th colspan="8">TBS DITERIMA PMKS</th>
		<th rowspan="2">RESTAN (JJG)</th>
	</tr>
	<tr>
		<th colspan="4">JANJANG PANEN</th>
		<th colspan="4">TONASE</th>
		<th colspan="4">JANJANG KIRIM</th>
		<th colspan="4">TONASE</th>
	</tr>
	<tr>
		<th colspan="2">BI</th>
		<th>BUDGET</th>
		<th rowspan="2">SBI</th>
		<th colspan="2">BI</th>
		<th>BUDGET</th>
		<th rowspan="2">SBI</th>
		<th colspan="2">BI</th>
		<th>BUDGET</th>
		<th rowspan="2">SBI</th>
		<th colspan="2">BI</th>
		<th>BUDGET</th>
		<th rowspan="2">SBI</th>
		<th rowspan="2">HI</th>
	</tr>
	<tr>
		<th>HI</th>
		<th>SHI</th>
		<th>BI</th>
		<th>HI</th>
		<th>SHI</th>
		<th>BI</th>
		<th>HI</th>
		<th>SHI</th>
		<th>BI</th>
		<th>HI</th>
		<th>SHI</th>
		<th>BI</th>
	</tr>
<%
Dim nsubhi, nsubshi, nsubbi, nsubsbi, nsubtonase_hi
Dim nsubtonase_shi, nsubtonase_bi, nsubtonase_sbi, nsubpom_hi, nsubpom_shi
Dim nsubpom_bi, nsubpom_sbi, nsubpom_tonase_hi, nsubpom_tonase_shi, nsubpom_tonase_bi, nsubpom_tonase_sbi, nsubrestan
	if not rsEPS.EOF then
		NO = 1
		nsubhi = 0
		nsubshi = 0
		nsubbi = 0
		nsubsbi = 0
		nsubtonase_hi = 0
		nsubtonase_shi = 0
		nsubtonase_bi = 0
		nsubtonase_sbi = 0
		nsubpom_hi = 0
		nsubpom_shi = 0
		nsubpom_bi = 0
		nsubpom_sbi = 0
		nsubpom_tonase_hi = 0
		nsubpom_tonase_shi = 0
		nsubpom_tonase_bi = 0
		nsubpom_tonase_sbi = 0
		nsubrestan = 0
		while not rsEPS.EOF
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center"><%=NO%></td>
				<td align="center"><%=rsEPS("fcba")%></td>
				<td align="center"><%=rsEPS("division")%></td>
				<td align="center"><%=rsEPS("plantingyear")%></td>
				<td align="center"><%=rsEPS("fieldcode")%></td>				
				<td align="right"><%=FormatNumber(rsEPS("hi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("shi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("bi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("sbi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("tonase_hi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("tonase_shi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("tonase_bi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("tonase_sbi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("pom_hi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("pom_shi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("pom_bi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("pom_sbi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("pom_tonase_hi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("pom_tonase_shi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("pom_tonase_bi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("pom_tonase_sbi"),2)%></td>
				<td align="right"><%=FormatNumber(rsEPS("restan"),2)%></td>
			</tr>
<%
			NO = NO + 1
			nsubhi = nsubhi + Cdbl(rsEPS("hi"))
			nsubshi = nsubshi + Cdbl(rsEPS("shi"))
			nsubbi = nsubbi + Cdbl(rsEPS("bi"))
			nsubsbi = nsubsbi + Cdbl(rsEPS("sbi"))
			nsubtonase_hi = nsubtonase_hi + Cdbl(rsEPS("tonase_hi"))
			nsubtonase_shi = nsubtonase_shi + Cdbl(rsEPS("tonase_shi"))
			nsubtonase_bi = nsubtonase_bi + Cdbl(rsEPS("tonase_bi"))
			nsubtonase_sbi = nsubtonase_sbi + Cdbl(rsEPS("tonase_sbi"))
			nsubpom_hi = nsubpom_hi + Cdbl(rsEPS("pom_hi"))
			nsubpom_shi = nsubpom_shi + Cdbl(rsEPS("pom_shi"))
			nsubpom_bi = nsubpom_bi + Cdbl(rsEPS("pom_bi"))
			nsubpom_sbi = nsubpom_sbi + Cdbl(rsEPS("pom_sbi"))
			nsubpom_tonase_hi = nsubpom_tonase_hi + Cdbl(rsEPS("pom_tonase_hi"))
			nsubpom_tonase_shi = nsubpom_tonase_shi + Cdbl(rsEPS("pom_tonase_shi"))
			nsubpom_tonase_bi = nsubpom_tonase_bi + Cdbl(rsEPS("pom_tonase_bi"))
			nsubpom_tonase_sbi = nsubpom_tonase_sbi + Cdbl(rsEPS("pom_tonase_sbi"))
			nsubrestan = nsubrestan + Cdbl(rsEPS("restan"))

			rsEPS.MoveNext			
		Wend
	end if
	rsEPS.Close
	set rsEPS = Nothing
%>
			<tr bgcolor="#9CB770">
				<td align="center" colspan="5">TOTAL</td>				
				<td align="right"><%=FormatNumber(nsubhi,2)%></td>
				<td align="right"><%=FormatNumber(nsubshi,2)%></td>
				<td align="right"><%=FormatNumber(nsubbi,2)%></td>
				<td align="right"><%=FormatNumber(nsubsbi,2)%></td>
				<td align="right"><%=FormatNumber(nsubtonase_hi,2)%></td>
				<td align="right"><%=FormatNumber(nsubtonase_shi,2)%></td>
				<td align="right"><%=FormatNumber(nsubtonase_bi,2)%></td>
				<td align="right"><%=FormatNumber(nsubtonase_sbi,2)%></td>
				<td align="right"><%=FormatNumber(nsubpom_hi,2)%></td>
				<td align="right"><%=FormatNumber(nsubpom_shi,2)%></td>
				<td align="right"><%=FormatNumber(nsubpom_bi,2)%></td>
				<td align="right"><%=FormatNumber(nsubpom_sbi,2)%></td>
				<td align="right"><%=FormatNumber(nsubpom_tonase_hi,2)%></td>
				<td align="right"><%=FormatNumber(nsubpom_tonase_shi,2)%></td>
				<td align="right"><%=FormatNumber(nsubpom_tonase_bi,2)%></td>
				<td align="right"><%=FormatNumber(nsubpom_tonase_sbi,2)%></td>
				<td align="right"><%=FormatNumber(nsubrestan,2)%></td>
			</tr>

</table>

<!--#include file="../Bottom_Report.asp"-->
