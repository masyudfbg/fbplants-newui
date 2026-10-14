<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%	
Dim rsSQL, strSQL
Dim v_txtBudgetYear, v_txtBudgetVersion, v_txtModuleVersion, v_txtyear, v_txtMonth, v_txttodaydate, tempDivision, tempPlantingDate, tempBlok, tempFcba, v_txtdatefrom, v_txtdateto
Dim tempStockGroup, tempStockCode, tempGroup, tempStatus, tempJobCode, tempVehicleGroup, tempVehicleCode,v_txtcomp
Dim GANGCODE, TempVarX


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

GANGCODE=request("GANGCODE")

v_txtcomp	= Request("label")
response.write v_txtcomp
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
<script language="JavaScript" type="text/javascript" >

function fireEvent(element,event){
    if (document.createEventObject){
    // dispatch for IE
    var evt = document.createEventObject();
    return element.fireEvent('on'+event,evt)
    }
    else{
    // dispatch for firefox + others
    var evt = document.createEvent("Events");
    evt.initEvent(event, true, true ); // event type,bubbling,cancelable
    return !element.dispatchEvent(evt);
    }
}
function set(){alert("asd");}

function fun(){    

	var CID = document.getElementById("combo1").value;	
	document.getElementById("compCode").value = CID;
	var StringAll;
	var Report001;
	var Report002;
	var Report003;
	var Report004;
	var Report005;
	var Report006;
	var Report007;
	var Report008;
	
    Report001 = "p_Summary_BA_NAF.asp?txtComp=" + CID + "&txtBudgetyear=<%=v_txtBudgetYear%>&txtBudgetVersion=<%=v_txtBudgetVersion%>&txtModuleVersion=<%=v_txtModuleVersion%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtdatefrom=<%=v_txttodaydate%>";
    Report001 = "<a href=" + Report001 +  " >Laporan Harian Kebun</a>";
	
    Report002 = "p_Summary_Production_NAF.asp?txtComp=" + CID + "&txtToDayDate=<%=v_txttodaydate%>" ;
	Report002 = "<a href=" + Report002 +  " >Laporan Harian Pabrik</a>";
	
	Report003 = "p_ProductionSummary_BA_NAF.asp?txtcomp=" + CID + "&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>" ;
	Report003 = "<a href=" + Report003 +  " >Produksi</a>";
	
	Report004 = "p_PerawatanTM_BA_NAF.asp?txtcomp=" + CID + "&txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtGroup=<%=tempGroup%>&txtStatusTmTbm=<%=tempStatus%>&txtJobCode=<%=tempJobCode%>" ;
	Report004 = "<a href=" + Report004 +  " >Perawatan</a>";
	
	Report005 = "p_StockGudangMonthly_BA_NAF.asp?txtcomp=" + CID + "&txtfcba=&txttodaydate=<%=v_txttodaydate%>&txtStockGroup=<%=tempStockGroup%>&txtStock=<%=tempStockCode%>";
	Report005 = "<a href=" + Report005 +  " >Stok Gudang</a>";
	
	Report006 = "p_AttendanceDaily_BA_NAF.asp?txtcomp=" + CID + "&txtDate=<%=v_txttodaydate%>&txtFcba=<%=tempFcba%>" ;
	Report006 = "<a href=" + Report006 +  " >Tenaga Kerja</a>";
	
	Report007 = "p_SolarDaily_BA_NAF.asp?txtcomp=" + CID + "&txtDate=<%=v_txttodaydate%>" ;
	Report007 = "<a href=" + Report007 +  " >Pemakaian Solar</a>";
	
	Report008 = "p_DailyVehicleSolar_BA_NAF.asp?txtcomp=" + CID + "&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>&txtfcba=<%=tempFcba%>";
	Report008 = "<a href=" + Report008 +  " >Rasio Pemakaian Solar</a> ";
						          
    document.getElementById("Report001").innerHTML=Report001;
    document.getElementById("Report002").innerHTML=Report002;
	document.getElementById("Report003").innerHTML=Report003;
    document.getElementById("Report004").innerHTML=Report004;
	document.getElementById("Report005").innerHTML=Report005;
    document.getElementById("Report006").innerHTML=Report006;
	document.getElementById("Report007").innerHTML=Report007;
    document.getElementById("Report008").innerHTML=Report008;

}
fun();
   
</script>

</head>


<body>
 
<form name="frmEntry" method="post" action="DashboardGroupManager.Asp">




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
			 <select name="txtBA" id="combo1" onchange="fun(this)">
			 <option value="ALL" <%if request.Form("txtBA")="ALL" then%> selected <%End if%>>::Please Select Company::</option>
			 <%
			 dim q
			 q="select companycode,companyname from COMPANY"
			 Set rs = connStr.Execute(q)
			 Do while not rs.EOF%>
			    <option value="<%=rs("companycode")%>"><%=rs("companyname")%></option>
			    <%rs.moveNext()
			 loop%>
			 rs.close()
			 rs=null%>
			</select>
	  </td>
	  <input type="text" id="compCode" />  
	  
		<div class="entry-header"><h2 class="headings">Laporan Harian</h2></div>
		<div class="entry-body">
			
			<div class="navlist">
				<h3></h3>
				<ul>
					<li>
					     <div id="Report001">
					         <a href="p_Summary_BA_NAF.asp?txtBudgetyear=<%=v_txtBudgetYear%>&txtBudgetVersion=<%=v_txtBudgetVersion%>&txtModuleVersion=<%=v_txtModuleVersion%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtdatefrom=<%=v_txttodaydate%>">Laporan Harian Kebun</a>
						 </div>
					</li>
					<li>
					     <div id="Report002">
					          <a href="p_Summary_Production_NAF.asp?txtToDayDate=<%=v_txttodaydate%>">Laporan Harian Pabrik</a>
						 </div>
					</li>
					     
					
					<li>
					    <div id="Report003">
							<a href="p_ProductionSummary_BA_NAF.asp?txtcomp=<%=TempVarX%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txttodaydate=<%=v_txttodaydate%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>">Produksi</a>
					    </div>
				    </li>
					<li>
					    <div id="Report004">
							<a href="p_PerawatanTM_BA_NAF.asp?txtfcba=<%=tempfcba%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtDivision=<%=tempDivision%>&txtBlok=<%=tempBlok%>&txtPlantingDate=<%=tempPlantingDate%>&txtGroup=<%=tempGroup%>&txtStatusTmTbm=<%=tempStatus%>&txtJobCode=<%=tempJobCode%>">Perawatan</a>
					     </div>
					</li>
					
					<li>
					     <div id="Report005">
							<a href="p_StockGudangMonthly_BA_NAF.asp?txtfcba=&txttodaydate=<%=v_txttodaydate%>&txtStockGroup=<%=tempStockGroup%>&txtStock=<%=tempStockCode%>">Stok Gudang</a>
					     </div>
					</li>
					<li>
					    <div id="Report006">
							<a href="p_AttendanceDaily_BA_NAF.asp?txtDate=<%=v_txttodaydate%>&txtFcba=<%=tempFcba%>">Tenaga Kerja</a>
				        </div>
				   </li>
				   <li>
					    <div id="Report007">
							<a href="p_SolarDaily_BA_NAF.asp?txtDate=<%=v_txttodaydate%>">Pemakaian Solar</a>
				        </div>
				   </li>					
					<li>
					   <div id="Report008">
							<a href="p_DailyVehicleSolar_BA_NAF.asp?txtdatefrom=<%=v_txtdatefrom%>&txtdateto=<%=v_txtdateto%>&txtVehicleGroup=<%=tempVehicleGroup%>&txtVehicleCode=<%=tempVehicleCode%>&txtfcba=<%=tempFcba%>">Rasio Pemakaian Solar</a>
					   </div>
					</li>

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

 </form>
</body>

  
	  
	 

</html>
