<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%	
Dim rsSQL, strSQL
Dim v_txtBudgetYear, v_txtBudgetVersion, v_txtModuleVersion, v_txtyear, v_txtMonth, v_txttodaydate, tempDivision, tempPlantingDate, tempBlok, tempFcba, v_txtdatefrom, v_txtdateto
Dim tempStockGroup, tempStockCode, tempGroup, tempStatus, tempJobCode, tempVehicleGroup, tempVehicleCode

strSQL = " select to_char(sysdate, 'YYYY') as CurrentYear, " & _
		 " to_char(sysdate, 'MM') as CurrentPeriod , " & _
		 " to_char(sysdate - 1, 'MM/DD/YYYY') as CurrentDate, " & _
         " to_date(to_char(sysdate - 1, 'MM')||'/01'|| to_char(sysdate - 1, 'YYYY'), 'MM/DD/YYYY') as StartMonthDate  from dual "

Set rsSQL = Server.CreateObject("ADODB.Recordset")

rsSQL.Open strSQL, connstr, adOpenStatic 

if not rsSQL.eof then
	v_txtBudgetyear = rsSQL("CurrentYear")
	v_txtBudgetVersion = 1
	v_txtModuleVersion = 1
	
	v_txtyear = rsSQL("CurrentYear")
	v_txtMonth = rsSQL("CurrentPeriod")
	v_txttodaydate = rsSQL("CurrentDate")
	v_txtdatefrom = rsSQL("StartMonthDate")
	v_txtdateto = rsSQL("CurrentDate")
End if
rsSQL.Close
set rsSQL = nothing

tempDivision = "--SUMMARY--" 
tempPlantingDate = "--SUMMARY--" 
tempBlok = "--SUMMARY--" 
tempFcba = "--SUMMARY--"

tempStockGroup = ""
tempStockCode = "--SUMMARY--"
tempGroup = ""
tempStatus = "--SUMMARY--"
tempJobCode = "--SUMMARY--"
tempVehicleGroup = ""
tempVehicleCode = "--SUMMARY--"
%>



<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta name="author" content="" />
<meta http-equiv="content-type" content="text/html;charset=iso-8859-1" />
<link rel="stylesheet" href="/css/stylereport.css" type="text/css" />
<title>ePlaS Reporting</title>


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
	  <td align="left"><font face="Arial" size="-1">Company Code :    </td>
	  <td align="left">
			 <select name="txtBA">
			 <%
			 dim q
			 q="select distinct fccompanycode,fccompanyname from BUSINESSUNIT"
			 Set rs = connStr.Execute(q)
			 Do while not rs.EOF%>
			    <option value="<%=rs("fccompanycode")%>"><%=rs("fccompanyname")%></option>
			    <%rs.moveNext()
			 loop%>
			 rs.close()
			 rs=null%>
			</select>
	  </td>
		<div class="entry-header"><h2 class="headings">Laporan Harian</h2></div>
		<div class="entry-body">
			
			<div class="navlist">
				<h3></h3>
				<ul>
					<li><a href="p_Summary_BA_NAF.asp?txtBudgetyear=<%=v_txtBudgetYear%>&txtBudgetVersion=<%=v_txtBudgetVersion%>&txtModuleVersion=<%=v_txtModuleVersion%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtdatefrom=<%=v_txttodaydate%>">Laporan Harian Kebun</a></li>
					<li><a href="p_Summary_Production_NAF.asp?txtToDayDate=<%=v_txttodaydate%>&txtFcba=<%=tempFcba%>">Laporan Harian Pabrik</a></li>
					
					<li><a href="p_ProductionSummary_BA_NAF.asp?txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">Produksi</a></li>
					
					<li><a href="p_PerawatanTM_BA_NAF.asp?txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtGroup=<%=tempGroup%>&txtStatusTmTbm=<%=tempStatus%>&txtJobCode=<%=tempJobCode%>">Perawatan</a></li>
					
					<!-- <li><a href="p_StockGudang_BA_NAF.asp">Stock Gudang</a></li> -->
					<li><a href="p_StockGudangMonthly_BA_NAF.asp?txtfcba=&txttodaydate=<%=v_txttodaydate%>&txtStockGroup=<%=tempStockGroup%>&txtStock=<%=tempStockCode%>">Stok Gudang</a></li>
					<li><a href="p_AttendanceDaily_BA_NAF.asp?txtDate=<%=v_txttodaydate%>&txtFcba=<%=tempFcba%>">Tenaga Kerja</a></li>
					<li><a href="p_SolarDaily_BA_NAF.asp?txtDate=<%=v_txttodaydate%>">Pemakaian Solar</a></li>					
					<li><a href="p_DailyVehicleSolar_BA_NAF.asp?txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>&txtfcba=<%=tempFcba%>">Rasio Pemakaian Solar</a></li>
					<li>&nbsp;</li>
				</ul>
			</div>
			
		</div><!--End Entry Body-->
		<div class='clear'></div>
	
		
	</div><!--End Content-->
	<div id="content">
		<div class="entry-header"><h2 class="headings">Laporan Bulanan</h2></div>
		<div class="entry-body">
			
			<div class="navlist">
				<h3></h3>
				<ul>
					<li><a href="p_Summary_BA_MONTH.asp?txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtdatefrom=<%=v_txttodaydate%>">Laporan Bulanan Kebun</a></li>
					<li><a href="Summary_Production_MONTH.asp?txtToDayDate=<%=v_txttodaydate%>&txtFcba=<%=tempFcba%>">Laporan Bulanan Pabrik</a></li>
					<li>&nbsp;</li>
				</ul>
			</div>
			
		</div><!--End Entry Body-->
		<div class='clear'></div>
	
		
	</div><!--End Content-->
	<div class='clear'></div>
</div><!--End Container-->
<div class='clear'></div>
</body>

  
	  
	 

</html>
