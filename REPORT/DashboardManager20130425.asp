<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%	
Dim rsSQL, strSQL,rsSQL1, strSQL1
Dim v_txtyear, v_txtMonth, v_txttodaydate, tempDivision, tempPlantingDate, tempBlok,txtComp

strSQL = " select to_char(sysdate, 'YYYY') as CurrentYear, " & _
		 " to_char(sysdate, 'MM') as CurrentPeriod , " & _
		 " to_char(sysdate, 'MM/DD/YYYY') as CurrentDate from dual "

Set rsSQL = Server.CreateObject("ADODB.Recordset")

rsSQL.Open strSQL, connstr, adOpenStatic 

if not rsSQL.eof then
	v_txtyear = rsSQL("CurrentYear")
	v_txtMonth = rsSQL("CurrentPeriod")
	v_txttodaydate = rsSQL("CurrentDate")
End if
rsSQL.Close


strSQL = "select fccompanycode  from businessunit where fccode = '" & Session("fcBA") & "'"
rsSQL.Open strSQL, connstr, adOpenStatic 
if not rsSQL.eof then
	txtComp = rsSQL("fccompanycode")
End if
rsSQL.Close

set rsSQL = nothing

tempDivision = "" 
tempPlantingDate = "--SUMMARY--" 
tempBlok = "--SUMMARY--" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="author" content="" />
<meta http-equiv="content-type" content="text/html;charset=iso-8859-1" />
<link rel="stylesheet" href="/css/stylereport.css" type="text/css" />
<title>iPS Reporting</title>
</head>
<body>
<div id="container">
	<div id="header" class="clearfix">
		<div class="left"><a href="../login.asp" title="Back to Home"><img src="../images/latar.jpg" alt="RW Plant Reports"/></a></div>
		<div class="right"><h1>Reports</h1></div>
		<div class='clear'></div>
	</div><!--End Header-->
	<div class='clear'></div>
	
	<div id="content">
		
		
		<div class="entry-header"><h2 class="headings">Daily Reports</h2></div>
		<div class="entry-body">
			<div class="navlist">
				<h3></h3>
				<ul>
					<li><a href="p_ProductionSummary_BA_NAF_M.asp?txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtComp=<%=txtComp%>">Production</a></li>								
					<li><a href="p_PerawatanTM_BA_NAF_M.asp?txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtComp=<%=txtComp%>">Perawatan</a></li>
					<li><a href="p_StockGudang_BA_NAF_M.asp">Stock Gudang</a></li>
					<li><a href="p_StockGudangMonthly_BA_NAF_M.asp">Stock Gudang per Period</a></li>
					<li><a href="p_AttendanceDaily_BA_NAF_M.asp">Tenaga Kerja</a></li>
					<li><a href="p_SolarDaily_BA_NAF_M.asp">Pemakaian Solar</a></li>					
					<li><a href="p_DailyVehicle_BA_NAF_M.asp">Output Kendaraan</a></li>
					<li><a href="p_DailyVehicleSolar_BA_NAF_M.asp">Fuel Consumption Kendaraan</a></li>	
				</ul>
			</div>
			
			
		</div><!--End Entry Body-->
		<div class='clear'></div>

		<!--div class="entry-header"><h2 class="headings">Monthly Reports</h2></div>
		<div class="entry-body">
			<div class="navlist">
				<h3></h3>
				<ul>
					<!--li><a href="p_Estate_Direct.Asp">Mature/Immature Cost Summary</a></li>
					<li><a href="p_CostElement.Asp?txtKind=Estate Cost Details">Mature/Immature Cost Details</a></li>
					<li><a href="p_inDirect.Asp?txtKind=DirectChargeMill">Indirect Cost</a></li>
					<li><a href="p_Mill_Direct.Asp?txtKind=in-Direct">Mill Cost</a></li-->
				</ul>
			</div>
			
		</div--><!--End Entry Body-->
		<div class='clear'></div>
	
		
	</div><!--End Content-->
	<div class='clear'></div>
</div><!--End Container-->
<div class='clear'></div>
</body>
</html>
