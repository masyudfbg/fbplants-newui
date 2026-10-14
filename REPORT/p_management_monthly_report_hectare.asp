<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtYear, v_txtMonth
Dim ScriptName
Dim rsMonthlyPupuk, strMonthlyPupuk, strMain, strMonthlyPupuk2
Dim cClass, nday, xNum

Dim var_colspan : var_colspan = 2

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType

v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")

Dim selectFields, strFieldCons, selectCons, whereCons, GroupbyCons, orderCons

strFieldCons	=	" USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantationcode" & _
					" , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantationcode" & _
					" , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregioncode" & _
					" , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) regioncode" & _
					" , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) companycode" & _
					" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba"
					
wherecons	=	" where companycode is not null "

If Len(Trim(v_txtGroupPlantation)) > 0 Then
	if v_txtGroupPlantation = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtGroupPlantation = "--ALL--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " And GroupPlantationcode = '" & v_txtGroupPlantation & "'"
		orderCons	=	orderCons	& " "
		var_colspan	=	var_colspan	+ 1
	End If
End If

If Len(Trim(v_txtPlantation)) > 0 Then
	if v_txtPlantation = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtPlantation = "--ALL--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Else
		selectCons	=	selectCons	& " "
		whereCons	=	whereCons	& " And Plantationcode = '" & v_txtPlantation & "'"
		orderCons	=	orderCons	& " "
	End If
End If

If Len(Trim(v_txtGroupRegion)) > 0 Then
	if v_txtGroupRegion = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtGroupRegion = "--ALL--" then
		selectCons	=	selectCons	& " , GroupRegioncode"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegioncode)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , GroupRegioncode"
		whereCons	=	whereCons	& " And GroupRegioncode = '" & v_txtGroupRegion & "'"
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegioncode)"
		var_colspan	=	var_colspan	+ 1
	End If
End If

If Len(Trim(v_txtRegion)) > 0 Then
	if v_txtRegion = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtRegion = "--ALL--" then
		selectCons	=	selectCons	& " , Regioncode"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Regioncode)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , Regioncode"
		whereCons	=	whereCons	& " And Regioncode = '" & v_txtRegion & "'"
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Regioncode)"
		var_colspan	=	var_colspan	+ 1
	End If
End If

If Len(Trim(v_txtCompany)) > 0 Then
	if v_txtCompany = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtCompany = "--ALL--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Else
		selectCons	=	selectCons	& " "
		whereCons	=	whereCons	& " And companycode = '" & v_txtCompany & "'"
		orderCons	=	orderCons	& " "
		
		q	=	" SELECT FCCODE" & _
				" , FCCOMPANYCODE" & _
				" , FCNAME" & _
				" , FCCOMPANYNAME" & _
				" , FCMANAGER" & _
				" , FCKTU" & _
				" , FCCROP" & _
				" , FCTYPE" & _
				" FROM BUSINESSUNIT WHERE" & _
				" FCCOMPANYCODE = '" & v_txtCompany & "'"	
		set rsCat = connStr.Execute(q)		
		if not rsCat.eof then
			CompanyName = rsCat("FCCOMPANYNAME")	
		end if
		q	=	""
		rsCat.Close
		set rsCat = Nothing
	End If
End If

If Len(Trim(v_txtfcba)) > 0 Then
	if v_txtfcba = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtfcba = "--ALL--" then
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , nourut"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	whereCons	& " And fcba = '" & v_txtfcba & "'"
		orderCons	=	orderCons	& " , nourut"
		var_colspan	=	var_colspan	+ 1
		
		q	=	" SELECT FCCODE" & _
				" , FCCOMPANYCODE" & _
				" , FCNAME" & _
				" , FCCOMPANYNAME" & _
				" , FCMANAGER" & _
				" , FCKTU" & _
				" , FCCROP" & _
				" , FCTYPE" & _
				" FROM BUSINESSUNIT WHERE" & _
				" FCCODE = '" & v_txtfcba & "'"	
		set rsCat = connStr.Execute(q)		
		if not rsCat.eof then
			fcbaname = rsCat("FCNAME")
			CompanyName=rsCat("FCCOMPANYNAME")
			MgrName = rsCat("FCMANAGER")
			KTUName = rsCat("FCKTU")
			CropName=rsCat("FCCROP")
			FcbaType=rsCat("FCTYPE")
		end if
		q	=	""
		rsCat.Close
		set rsCat = Nothing
	End If
End If

v_txtYear	=	Request("txtYear")
v_txtMonth	=	Request("txtMonth")

If Len(Trim(v_txtYear)) = 0 Then
	v_txtYear = Year(Session("v_Costbook"))
End If

If Len(Trim(v_txtMonth)) = 0 Then
	v_txtMonth = Month(Session("v_Costbook"))
End If

Dim varLastDayofMonth : varLastDayofMonth = v_txtMonth & "/" & GetLastDay(v_txtMonth,v_txtYear) & "/" & v_txtYear

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain	=	" SELECT" & strFieldCons & _			
			"  ,NVL (sum(mature),0) mature" & _
			" , NVL (sum(bud_mature),0) bud_mature" & _
			" , NVL (sum(immature_01),0) immature_01" & _
		    " , NVL (sum(immature_02),0) immature_02" & _
		    " , NVL (sum(immature_03),0) immature_03" & _
			" , NVL (sum(bud_immature),0) bud_immature" & _
		    " , NVL (sum(newplanting_thismonth),0) newplanting_thismonth" & _
		    " , NVL (sum(newplanting_todate),0) newplanting_todate" & _
		    " , NVL (sum(budget_annual),0) budget_annual " & _
		    " , NVL (sum(budget_thismonth),0) budget_thismonth" &_
		    " , NVL (sum(budget_todate),0) budget_todate" &_
			"     FROM (    " & _
			"         SELECT * " & _
			"         FROM" & _
			"             MGT2_HECTARE_DET_TAB" & _
			"         WHERE" & _
			"             YEAR = " & v_txtYear & "" & _
			"             AND MONTH = " & v_txtMonth & "" & _
			"         )" & _
			"     GROUP BY" & _
			"         FCBA"								
strMonthlyPupuk2	=    " , NVL (sum(mature),0) mature" & _
						 " , NVL (sum(bud_mature),0) bud_mature" & _
						 " , NVL (sum(immature_01),0) immature_01" & _
						 " , NVL (sum(immature_02),0) immature_02" & _
						 " , NVL (sum(immature_03),0) immature_03" & _						 
						 " , NVL (sum(bud_immature),0) bud_immature" & _
						 " , NVL (sum(newplanting_thismonth),0) newplanting_thismonth" & _
						 " , NVL (sum(newplanting_todate),0) newplanting_todate" & _
						 " , NVL (sum(budget_annual),0) budget_annual " & _
						 " , NVL (sum(budget_thismonth),0) budget_thismonth" &_
						 " , NVL (sum(budget_todate),0) budget_todate"
  

strMonthlyPupuk	=	" select GroupPlantationcode, Plantationcode, Companycode" & selectCons & strMonthlyPupuk2 &"  From (" & strMain & " )" & _
					whereCons & _
					" group by GroupPlantationcode, Plantationcode, Companycode" & selectCons & "" & _
					" Order by GroupPlantationcode,Plantationcode, Companycode" & orderCons & ""

'response.write strMonthlyPupuk
%>
<html>
<head>
<title>ePlaS Reporting - Bulanan Areal</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
 <script>
$(function() {
$( ".datepicker" ).datepicker();
});
</script>
<!--<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>-->

<style type="text/css">
@media print
{
.printhide{display:none;}
}
</style>
</head>	
<body>
	<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="menutop.asp"-->
			</div>
			<form action="<%=ScriptName%>" method="post" name="frmEntry">
			<div class="printhide">
				<!--#include virtual="/REPORT/FormSelectNewSumE.html"-->
				<table id="form-table" class="centered-margin">
					<tr>
						<td>
							<font face="verdana" size="1">TAHUN</font>
						</td>
						<% 
							Dim mon, Yr, MinYr, MaxYr, i, MM
							mon = Month(Session("v_Costbook"))
							If len(v_txtMonth)=0 then v_txtMonth=mon
						
							Yr = Year(Session("v_Costbook"))
							MinYr = 2015
							MaxYr = Yr  + 2
							If len(v_txtYear)= 0 then v_txtYear = Yr
						%>
						<td width="4">
							<select name="txtYear">
							<%For i = MinYr to MaxYr%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
							<%Next%>
							</select>
						</td>
						<td>
							<font face="verdana" size="1">BULAN</font>
						</td>
						<td width="2">
							<select name="txtMonth">
							<%For i = 1 to 12%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
							<%Next%>
							</select>
						</td>
						<td>
							<input type="submit" value="View" />
						</td>
					</tr>
				</table>
			</div>
			</form>
		</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead style="border-left:1px #000000 solid;">
	<tr>
         <th colspan="<%=var_colspan+13%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
            <div id="header">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
            </div>
            <div id="subheader" class="centered">
                <h2>LAPORAN BULANAN AREAL STATEMENT</h3>
                <br/>
                <h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
            </div>
		 </th>
	</tr>
	<tr>
		<th rowspan="3" style="background-color:#00A300;color:#fff;">NO</th>		
		<%If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleGroupPlantation%></th>
		<%End If%>
		<%If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
		<%End If%>
		<%If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
		<%End If%>
		<th rowspan="3" style="background-color:#00A300;color:#fff;">PT</th>
		<%If Len(Trim(v_txtfcba)) > 0 And v_txtfcba <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
		<%End If%>
		<th rowspan="3">MATURE</th>
		<th rowspan="3">BUDGET MATURE</th>
		<th colspan="4">Immature</th>
		<th rowspan="3">BUDGET IMMATURE</th>
		<th colspan="5">New Planting</th>
		<th style="border-bottom:none;">TOTAL</th>		
	</tr>
	<tr>
		<th rowspan="2">1</th>
        <th rowspan="2">2</th>
		<th rowspan="2">3</th>
		<th rowspan="2" style="border-bottom:none;">Total</th>
		<th colspan="2">This Month</th>
		<th colspan="2">To Date</th>
        <th colspan="1">Annual</th>
		<th style=" border-top:none; border-bottom:none;" >Planted</th>
	</tr>
	<tr>
		<th>Act</th>
		<th>Bgt</th>
		<th>Act</th>
		<th>Bgt</th>
        <th>Budget</th>
        <th style=" border-top:none;">Area</th>
    </tr>
    </thead>
	<tbody>
<%
	Set rsMonthlyPupuk = Server.CreateObject("ADODB.Recordset")
	rsMonthlyPupuk.Open strMonthlyPupuk, connstr, adOpenStatic
	
	Dim Ttl_MATURE, Ttl_ACT_THISMONTH,Ttl_BUD_THISMONTH, Ttl_THISMONTHPER, Ttl_ACT_TODATE,Ttl_BUD_TODATE, Ttl_TODATEPER, Ttl_BUD_ANNUAL, TTL_BUD_IMMATURE, TTL_BUD_MATURE
	Dim GTtl_MATURE, GTTL_IMMATURE_01, GTTL_IMMATURE_02, GTTL_IMMATURE_03, GTTL_TOTIMATURE, GTTL_NEWPLANTING_THISMONTH, GTTL_BUDGET_THISMONTH, GTTL_NEWPLANTING_TODATE, GTTL_BUD_IMMATURE, GTTL_BUD_MATURE
	Dim GTTL_BUDGET_TODATE, GTTL_BUDGET_ANNUAL, GTTL_TOTPLANTED
	
	Dim tempGroupPlantation : tempGroupPlantation = ""
	Dim tempPlantation : tempPlantation = ""
	Dim tempGroupRegion : tempGroupRegion = ""
	Dim tempRegion : tempRegion = ""
	Dim tempCompany : tempCompany = ""
	Dim tempfcba : tempfcba = ""
	Dim tempLocationCode : tempLocationCode = ""

	Dim SprevGroupPlantation : SprevGroupPlantation = ""
	Dim SprevPlantation : SprevPlantation = ""
	Dim SprevGroupRegion : SprevGroupRegion = ""
	Dim SprevRegion : SprevRegion = ""
	Dim SprevCompany : SprevCompany = ""
	Dim sprevFcba : sprevFcba = ""
	Dim sprevLocationCode : sprevLocationCode = ""

	if rsMonthlyPupuk.bof and rsMonthlyPupuk.eof then
%>
	<tr>
		<td colspan="<%=var_colspan+10%>">NO DATA AVAILABLE</td>
	</tr>
<%
	else
		Dim TOTIMATURE,MATURE,NEWPLANTING_TODATE, immature_01, immature_02, immature_03, TOTPLANTED
		Dim TTL_IMMATURE_01, TTL_IMMATURE_02, TTL_IMMATURE_03, TTL_TOTIMATURE, TTL_NEWPLANTING_THISMONTH, TTL_BUDGET_THISMONTH
		Dim TTL_NEWPLANTING_TODATE, TTL_BUDGET_TODATE, TTL_BUDGET_ANNUAL, TTL_TOTPLANTED
		Dim NOX : NOX = 1
		while not rsMonthlyPupuk.eof
%>
	<% 
	Dim strRowsCount,rsSqlRow
		strRowsCount	= "Select Count(Plantationcode)RowsCount FROM(" & strMonthlyPupuk & ") Where Plantationcode = '"& rsMonthlyPupuk("Plantationcode") &"'"
		Set rsSqlRow = Server.CreateObject("ADODB.Recordset")
		rsSqlRow.Open strRowsCount ,connstr, adOpenStatic
		
			MATURE=rsMonthlyPupuk("MATURE")
			NEWPLANTING_TODATE = rsMonthlyPupuk("NEWPLANTING_TODATE")
			immature_01 = rsMonthlyPupuk("immature_01")
			immature_02 = rsMonthlyPupuk("immature_02")
			immature_03 = rsMonthlyPupuk("immature_03")
				
			 TOTIMATURE = cdbl(immature_01) + cdbl(immature_02) + cdbl(immature_03)
			 TOTPLANTED = cdbl(MATURE) + cdbl(NEWPLANTING_TODATE) + cdbl(TOTIMATURE) 
		
		if NOX = 1 then
			Dim IDPlantation,IDPlantation2  
			IF rsMonthlyPupuk("Plantationcode") = "PLT1" OR rsMonthlyPupuk("Plantationcode") = "1(SATU)" then
				IDPlantation = "A"
				IDPlantation2 = "Plantation-1"
			Elseif rsMonthlyPupuk("Plantationcode") = "PLT2" OR rsMonthlyPupuk("Plantationcode") = "2(DUA)" then
				IDPlantation = "B"
				IDPlantation2 = "Plantation-2"
			Else 
				IDPlantation = "X"
				IDPlantation2 = "Plantation-X"
			End iF
		%><tr>
			<td align="Center" ><b><%=IDPlantation%></b></td>
			<td align="Left" colspan="<%=var_colspan+12	%>"><b><%=IDPlantation2%></b></td>
		</tr>
		<%
		End If
		%>
			<tr>
			<td align="right">&nbsp;</td>
		<%
		if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% 
					   tempGroupPlantation = rsMonthlyPupuk("GroupPlantationcode")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
							<%=rsMonthlyPupuk("GroupPlantationcode")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
		<% end if 
		if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% 					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsMonthlyPupuk("GroupPlantationcode")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsMonthlyPupuk("Plantationcode")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsMonthlyPupuk("GroupRegioncode")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempLocationCode = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="Left">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
							<%=rsMonthlyPupuk("GroupRegioncode")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="Left"></td>					
					<% end if %>	
				<% end if %>

		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% 
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsMonthlyPupuk("GroupPlantationcode")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsMonthlyPupuk("Plantationcode")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--"  and  v_txtGroupRegion <> "" then
					       tempGroupRegion = rsMonthlyPupuk("GroupRegioncode")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsMonthlyPupuk("Regioncode") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempLocationCode = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="Left">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
							<%=rsMonthlyPupuk("regioncode")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="Left"></td>					
					<% end if %>	
				<% end if %>
				
				<%
					Dim tmpCompany
					tmpCompany = rsMonthlyPupuk("Companycode")
					If tmpCompany <> SprevCompany then
				%>
					<td align="Left"><%=rsMonthlyPupuk("Companycode")%></td>
				<%
					SprevCompany = tmpCompany
					else 
				%>
					<td align="Left"></td>
				<% 
					End if
			
			if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsMonthlyPupuk("GroupPlantationcode")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsMonthlyPupuk("Plantationcode")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion  <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsMonthlyPupuk("GroupRegioncode")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion  <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsMonthlyPupuk("Regioncode") 
					   else
					       tempRegion = ""
					   end if
					   
					   if  v_txtcompany  <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsMonthlyPupuk("Companycode")
					   else
					       tempCompany =""
					   end if
					   
					   tempfcba = rsMonthlyPupuk("fcba")
					   tempLocationCode = "--SUMMARY--"
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="Left">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
						    <%=rsMonthlyPupuk("fcba")%>
							</a>
						</td>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="Left"></td>					
					<% end if %>
				<% end if %>
		<td align="right"><%=FormatNumber(rsMonthlyPupuk("MATURE"),0)%></td>
		<td align="right"><%=FormatNumber(rsMonthlyPupuk("bud_mature"),0)%></td>
		<td align="right"><%=FormatNumber(rsMonthlyPupuk("IMMATURE_01"),0)%></td>		
		<td align="right"><%=FormatNumber(rsMonthlyPupuk("IMMATURE_02"),0)%></td>	
		<td align="right"><%=FormatNumber(rsMonthlyPupuk("IMMATURE_03"),0)%></td>
		<td align="right"><%=FormatNumber(TOTIMATURE,0)%></td>
		<td align="right"><%=FormatNumber(rsMonthlyPupuk("bud_immature"),0)%></td>
		<td align="right"><%=FormatNumber(rsMonthlyPupuk("NEWPLANTING_THISMONTH"),0)%></td>
		<td align="right"><%=FormatNumber(rsMonthlyPupuk("BUDGET_THISMONTH"),0)%></td>
		<td align="right"><%=FormatNumber(rsMonthlyPupuk("NEWPLANTING_TODATE"),0)%></td>
		<td align="right"><%=FormatNumber(rsMonthlyPupuk("BUDGET_TODATE"),0)%></td>
		<td align="right"><%=FormatNumber(rsMonthlyPupuk("BUDGET_ANNUAL"),0)%></td>
		<td align="right"><%=FormatNumber(TOTPLANTED,0)%></td>
	</tr>
<%
		Ttl_MATURE               	= Ttl_MATURE + Cdbl(rsMonthlyPupuk("MATURE"))
		TTL_BUD_MATURE				= TTL_BUD_MATURE + Cdbl(rsMonthlyPupuk("bud_mature"))
		TTL_IMMATURE_01				= TTL_IMMATURE_01 + Cdbl(rsMonthlyPupuk("IMMATURE_01")) 
		TTL_IMMATURE_02 			= TTL_IMMATURE_02 + Cdbl(rsMonthlyPupuk("IMMATURE_02"))
		TTL_IMMATURE_03 			= TTL_IMMATURE_03 + Cdbl(rsMonthlyPupuk("IMMATURE_03"))
		TTL_TOTIMATURE 	    		= TTL_TOTIMATURE + TOTIMATURE
		TTL_BUD_IMMATURE			= TTL_BUD_IMMATURE + Cdbl(rsMonthlyPupuk("bud_immature"))
		TTL_NEWPLANTING_THISMONTH 	= TTL_NEWPLANTING_THISMONTH + Cdbl(rsMonthlyPupuk("NEWPLANTING_THISMONTH"))
		TTL_BUDGET_THISMONTH 		= TTL_BUDGET_THISMONTH + Cdbl(rsMonthlyPupuk("BUDGET_THISMONTH"))
		TTL_NEWPLANTING_TODATE		= TTL_NEWPLANTING_TODATE + Cdbl(rsMonthlyPupuk("NEWPLANTING_TODATE"))
		TTL_BUDGET_TODATE			= TTL_BUDGET_TODATE + Cdbl(rsMonthlyPupuk("BUDGET_TODATE"))
		TTL_BUDGET_ANNUAL			= TTL_BUDGET_ANNUAL + Cdbl(rsMonthlyPupuk("BUDGET_ANNUAL"))
		TTL_TOTPLANTED				= TTL_TOTPLANTED + TOTPLANTED
		
		if NOX = CDbl(rsSqlRow("RowsCount"))then
			NOX = 1
		%>
		<tr>
			<td align="Center" style="background-color:#FFFF00;" colspan="<%=var_colspan%>"><b>Total</b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Ttl_MATURE,0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(TTL_BUD_MATURE,0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(TTL_IMMATURE_01,0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(TTL_IMMATURE_02,0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(TTL_IMMATURE_03,0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(TTL_TOTIMATURE,0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(TTL_BUD_IMMATURE,0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(TTL_NEWPLANTING_THISMONTH,0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(TTL_BUDGET_THISMONTH,0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(TTL_NEWPLANTING_TODATE,0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(TTL_BUDGET_TODATE,0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(TTL_BUDGET_ANNUAL,0)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(TTL_TOTPLANTED,0)%></b></td>
		</tr>
		
		<%
			Ttl_MATURE					= 0
			TTL_BUD_MATURE				= 0
			TTL_IMMATURE_01 			= 0
			TTL_IMMATURE_02 			= 0
			TTL_IMMATURE_03 			= 0
			TTL_TOTIMATURE 	    		= 0
			TTL_BUD_IMMATURE			= 0
			TTL_NEWPLANTING_THISMONTH 	= 0
			TTL_BUDGET_THISMONTH 		= 0
			TTL_NEWPLANTING_TODATE		= 0
			TTL_BUDGET_TODATE			= 0
			TTL_BUDGET_ANNUAL			= 0
			TTL_TOTPLANTED				= 0
			
			
			
			
		Else
			NOX = NOX+ 1
		End IF
		
		
				GTtl_MATURE 				= GTtl_MATURE + Cdbl(rsMonthlyPupuk("MATURE"))
				GTTL_BUD_MATURE				= GTTL_BUD_MATURE + Cdbl(rsMonthlyPupuk("BUD_MATURE"))
				GTTL_IMMATURE_01 			= GTTL_IMMATURE_01 + Cdbl(rsMonthlyPupuk("IMMATURE_01"))
				GTTL_IMMATURE_02 			= GTTL_IMMATURE_02 + Cdbl(rsMonthlyPupuk("IMMATURE_02"))
				GTTL_IMMATURE_03 			= GTTL_IMMATURE_03 + Cdbl(rsMonthlyPupuk("IMMATURE_03"))
				GTTL_TOTIMATURE 	    	= GTTL_TOTIMATURE + TOTIMATURE
				GTTL_BUD_IMMATURE			= GTTL_BUD_IMMATURE + Cdbl(rsMonthlyPupuk("BUD_IMMATURE"))
				GTTL_NEWPLANTING_THISMONTH 	= GTTL_NEWPLANTING_THISMONTH + Cdbl(rsMonthlyPupuk("NEWPLANTING_THISMONTH"))
				GTTL_BUDGET_THISMONTH 		= GTTL_BUDGET_THISMONTH + Cdbl(rsMonthlyPupuk("BUDGET_THISMONTH"))
				GTTL_NEWPLANTING_TODATE		= GTTL_NEWPLANTING_TODATE + Cdbl(rsMonthlyPupuk("NEWPLANTING_TODATE"))
				GTTL_BUDGET_TODATE			= GTTL_BUDGET_TODATE + Cdbl(rsMonthlyPupuk("BUDGET_TODATE"))
				GTTL_BUDGET_ANNUAL			= GTTL_BUDGET_ANNUAL + Cdbl(rsMonthlyPupuk("BUDGET_ANNUAL"))
				GTTL_TOTPLANTED				= GTTL_TOTPLANTED + TOTPLANTED
		
			
		rsMonthlyPupuk.MoveNext
		wend
		
			
			%>
			<tr>
				<td align="Center" style="background-color:#FFFF00;" colspan="<%=var_colspan%>"><b>Grand Total</b></td>
				<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(GTtl_MATURE,0)%></b></td>
				<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(GTTL_BUD_MATURE ,0)%></b></td>
				<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(GTTL_IMMATURE_01,0)%></b></td>
				<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(GTTL_IMMATURE_02,0)%></b></td>
				<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(GTTL_IMMATURE_03,0)%></b></td>
				<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(GTTL_TOTIMATURE,0)%></b></td>
				<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(GTTL_BUD_IMMATURE,0)%></b></td> 
				<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(GTTL_NEWPLANTING_THISMONTH,0)%></b></td>
				<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(GTTL_BUDGET_THISMONTH,0)%></b></td>
				<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(GTTL_NEWPLANTING_TODATE,0)%></b></td>
				<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(GTTL_BUDGET_TODATE,0)%></b></td>
				<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(GTTL_BUDGET_ANNUAL,0)%></b></td>
				<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(GTTL_TOTPLANTED,0)%></b></td>
			</tr>
			<%
				GTtl_MATURE					= 0
				GTTL_BUD_MATURE				= 0
				GTTL_IMMATURE_01 			= 0
				GTTL_IMMATURE_02 			= 0
				GTTL_IMMATURE_03 			= 0
				GTTL_TOTIMATURE 	    	= 0
				GTTL_BUD_IMMATURE 			= 0
				GTTL_NEWPLANTING_THISMONTH 	= 0
				GTTL_BUDGET_THISMONTH 		= 0
				GTTL_NEWPLANTING_TODATE		= 0
				GTTL_BUDGET_TODATE			= 0
				GTTL_BUDGET_ANNUAL			= 0
				GTTL_TOTPLANTED				= 0
			
	rsMonthlyPupuk.Close
	Set rsMonthlyPupuk = Nothing
	end if
%>
</tbody>
</table>
</body>
</html>