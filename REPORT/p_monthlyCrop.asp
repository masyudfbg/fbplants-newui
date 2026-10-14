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

strFieldCons	=	" USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation" & _
					" , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation" & _
					" , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion" & _
					" , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region" & _
					" , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company" & _
					" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba"
					
wherecons	=	" where company is not null "

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
		whereCons	=	wherecons	& " And GroupPlantation = '" & v_txtGroupPlantation & "'"
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
		whereCons	=	whereCons	& " And Plantation = '" & v_txtPlantation & "'"
		orderCons	=	orderCons	& " "
	End If
End If

If Len(Trim(v_txtGroupRegion)) > 0 Then
	if v_txtGroupRegion = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtGroupRegion = "--ALL--" then
		selectCons	=	selectCons	& " , GroupRegion"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , GroupRegion"
		whereCons	=	whereCons	& " And GroupRegion = '" & v_txtGroupRegion & "'"
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_colspan	=	var_colspan	+ 1
	End If
End If

If Len(Trim(v_txtRegion)) > 0 Then
	if v_txtRegion = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtRegion = "--ALL--" then
		selectCons	=	selectCons	& " , Region"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , Region"
		whereCons	=	whereCons	& " And Region = '" & v_txtRegion & "'"
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
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
		whereCons	=	whereCons	& " And company = '" & v_txtCompany & "'"
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
			" ,NVL(SUM (ACT_THISMONTH),0)ACT_THISMONTH"&_
			" ,NVL(SUM (BUD_THISMONTH),0)BUD_THISMONTH"&_
			" ,NVL(SUM (ACT_TODATE),0)ACT_TODATE"&_ 
			" ,NVL(SUM (BUD_TODATE),0)BUD_TODATE"&_
			" ,NVL(SUM (BUD_ANNUAL),0)BUD_ANNUAL"  &_
			"     FROM (    " & _
			"         SELECT * " & _
			"         FROM" & _
			"             MGT2_PRODUCTION_DET_TAB" & _
			"         WHERE" & _
			"             YEAR = " & v_txtYear & "" & _
			"             AND MONTH <= " & v_txtMonth & "" & _
			"         )" & _
			"     GROUP BY" & _
			"         FCBA"								
strMonthlyPupuk2	=	" ,SUM (ACT_THISMONTH)ACT_THISMONTH"&_
						" ,SUM(BUD_THISMONTH)BUD_THISMONTH"&_
						" ,SUM(ACT_TODATE)ACT_TODATE"&_ 
						" ,SUM (BUD_TODATE)BUD_TODATE"&_
						" ,SUM (BUD_ANNUAL)BUD_ANNUAL"  

strMonthlyPupuk	=	" select GroupPlantation, Plantation, Company" & selectCons & strMonthlyPupuk2 &"  From (" & strMain & " )" & _
					whereCons & _
					" group by GroupPlantation, Plantation, Company" & selectCons & "" & _
					" Order by GroupPlantation, Plantation, Company" & orderCons & ""
'response.write strMonthlyPupuk
%>
<html>
<head>
<title>ePlaS Reporting - Bulanan Crop</title>
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
         <th colspan="<%=var_colspan+8%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
            <div id="header">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
            </div>
            <div id="subheader" class="centered">
                <h2>LAPORAN BULANAN CROP</h3>
                <br/>
                <h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
            </div>
		 </th>
	</tr>
	<tr>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">NO</th>
		<%If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleGroupPlantation%></th>
		<%End If%>
		<%If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
		<%End If%>
		<%If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
		<%End If%>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">PT</th>
		<%If Len(Trim(v_txtfcba)) > 0 And v_txtfcba <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
		<%End If%>
		<th colspan="3" style="background-color:#00A300;color:#fff;">THIS MONTH</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">TO DATE</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">ANNUAL BUDGET</th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;">ACT</th>
		<th style="background-color:#00A300;color:#fff;">BGT</th>
		<th style="background-color:#00A300;color:#fff;">%</th>
		<th style="background-color:#00A300;color:#fff;">ACT</th>
		<th style="background-color:#00A300;color:#fff;">BGT</th>
		<th style="background-color:#00A300;color:#fff;">%</th>
	</tr>
    </thead>
	<tbody>
<%
	Set rsMonthlyPupuk = Server.CreateObject("ADODB.Recordset")
	rsMonthlyPupuk.Open strMonthlyPupuk, connstr, adOpenStatic
	
	Dim Ttl_ACT_THISMONTH,Ttl_BUD_THISMONTH, Ttl_THISMONTHPER, Ttl_ACT_TODATE,Ttl_BUD_TODATE, Ttl_TODATEPER, Ttl_BUD_ANNUAL
	
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
		<td colspan="<%=var_colspan+8%>">NO DATA AVAILABLE</td>
	</tr>
<%
	else
		Dim NOX : NOX = 1
		while not rsMonthlyPupuk.eof
%>
	<% 
	Dim strRowsCount,rsSqlRow
		strRowsCount	= "Select Count(Plantation)RowsCount FROM(" & strMonthlyPupuk & ") Where Plantation = '"& rsMonthlyPupuk("Plantation") &"'"
		Set rsSqlRow = Server.CreateObject("ADODB.Recordset")
		rsSqlRow.Open strRowsCount ,connstr, adOpenStatic
		
		if NOX = 1 then
			Dim IDPlantation,IDPlantation2  
			IF rsMonthlyPupuk("Plantation") = "PLT1" OR rsMonthlyPupuk("Plantation") = "1(SATU)" then
				IDPlantation = "A"
				IDPlantation2 = "Plantation-1"
			Elseif rsMonthlyPupuk("Plantation") = "PLT2" OR rsMonthlyPupuk("Plantation") = "2(DUA)" then
				IDPlantation = "B"
				IDPlantation2 = "Plantation-2"
			Else 
				IDPlantation = "X"
				IDPlantation2 = "Plantation-X"
			End iF
		%><tr>
			<td align="Center" ><b><%=IDPlantation%></b></td>
			<td align="Left" colspan="<%=var_colspan+7%>"><b><%=IDPlantation2%></b></td>
		</tr>
		<%
		End If
		%>
			<tr>
			<td align="right">&nbsp;</td>
		<%
		if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% 
					   tempGroupPlantation = rsMonthlyPupuk("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
							<%=rsMonthlyPupuk("GroupPlantation")%>
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
					      tempGroupPlantation = rsMonthlyPupuk("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsMonthlyPupuk("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsMonthlyPupuk("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempLocationCode = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="Left">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
							<%=rsMonthlyPupuk("GroupRegion")%>
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
					      tempGroupPlantation = rsMonthlyPupuk("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsMonthlyPupuk("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--"  and  v_txtGroupRegion <> "" then
					       tempGroupRegion = rsMonthlyPupuk("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsMonthlyPupuk("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempLocationCode = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="Left">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>">
							<%=rsMonthlyPupuk("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="Left"></td>					
					<% end if %>	
				<% end if %>
				
				<%
					Dim tmpCompany
					tmpCompany = rsMonthlyPupuk("Company")
					If tmpCompany <> SprevCompany then
				%>
					<td align="Left"><%=rsMonthlyPupuk("Company")%></td>
				<%
					SprevCompany = tmpCompany
					else 
				%>
					<td align="Left"></td>
				<% 
					End if
			
			if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsMonthlyPupuk("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsMonthlyPupuk("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion  <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsMonthlyPupuk("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion  <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsMonthlyPupuk("Region") 
					   else
					       tempRegion = ""
					   end if
					   
					   if  v_txtcompany  <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsMonthlyPupuk("Company")
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
		<td align="right"><%=FormatNumber(rsMonthlyPupuk("ACT_THISMONTH"),2)%></td>
		<td align="right"><%=FormatNumber(rsMonthlyPupuk("BUD_THISMONTH"),2)%></td>
		<%
			Dim THISMONTHPER
			if Cdbl(rsMonthlyPupuk("BUD_THISMONTH")) <> 0 then
				THISMONTHPER = Cdbl(rsMonthlyPupuk("ACT_THISMONTH")) / Cdbl(rsMonthlyPupuk("BUD_THISMONTH")) * 100
			else
				THISMONTHPER = 0
			End if
		%>
		<td align="right"><%=FormatNumber(THISMONTHPER,2)%></td>
		<td align="right"><%=FormatNumber(rsMonthlyPupuk("ACT_TODATE"),2)%></td>
		<td align="right"><%=FormatNumber(rsMonthlyPupuk("BUD_TODATE"),2)%></td>
		<%
			Dim TODATEPER
			if Cdbl(rsMonthlyPupuk("BUD_TODATE")) <> 0 then
				TODATEPER = Cdbl(rsMonthlyPupuk("ACT_TODATE")) / Cdbl(rsMonthlyPupuk("BUD_TODATE")) * 100
			else
				TODATEPER = 0
			End if
		%>
		<td align="right"><%=FormatNumber(TODATEPER,2)%></td>
		<td align="right"><%=FormatNumber(rsMonthlyPupuk("BUD_ANNUAL"),2)%></td>
	</tr>
<%
		Ttl_ACT_THISMONTH 	= Ttl_ACT_THISMONTH + Cdbl(rsMonthlyPupuk("ACT_THISMONTH"))
		Ttl_BUD_THISMONTH 	= Ttl_BUD_THISMONTH + Cdbl(rsMonthlyPupuk("BUD_THISMONTH"))
		Ttl_ACT_TODATE 		= Ttl_ACT_TODATE + Cdbl(rsMonthlyPupuk("ACT_TODATE"))
		Ttl_BUD_TODATE 		= Ttl_BUD_TODATE + Cdbl(rsMonthlyPupuk("BUD_TODATE"))
		Ttl_BUD_ANNUAL		= Ttl_BUD_ANNUAL + Cdbl(rsMonthlyPupuk("BUD_ANNUAL"))
		
		if NOX = CDbl(rsSqlRow("RowsCount"))then
			NOX = 1
		%>
		<tr>
			<td align="Center" style="background-color:#FFD700;" colspan="<%=var_colspan%>"><b>Total</b></td>
			<td align="right" style="background-color:#FFD700;"><b><%=FormatNumber(Ttl_ACT_THISMONTH,2)%></b></td>
			<td align="right" style="background-color:#FFD700;"><b><%=FormatNumber(Ttl_BUD_THISMONTH,2)%></b></td>
			<%
			if Cdbl(Ttl_BUD_THISMONTH) <> 0 then
				Ttl_THISMONTHPER 	= Ttl_ACT_THISMONTH / Ttl_BUD_THISMONTH * 100
			Else 
				Ttl_THISMONTHPER	= 0
			End if
			%>
			<td align="right" style="background-color:#FFD700;"><b><%=FormatNumber(Ttl_THISMONTHPER,2)%></b></td>
			<td align="right" style="background-color:#FFD700;"><b><%=FormatNumber(Ttl_ACT_TODATE,2)%></b></td>
			<td align="right" style="background-color:#FFD700;"><b><%=FormatNumber(Ttl_BUD_TODATE,2)%></b></td>
			<%
			if Cdbl(Ttl_BUD_TODATE) <> 0 then
				Ttl_TODATEPER 	= Ttl_ACT_TODATE / Ttl_BUD_TODATE * 100
			Else 
				Ttl_TODATEPER	= 0
			End if
			%>
			<td align="right" style="background-color:#FFD700;"><b><%=FormatNumber(Ttl_TODATEPER,2)%></b></td>
			<td align="right" style="background-color:#FFD700;"><b><%=FormatNumber(Ttl_BUD_ANNUAL,2)%></b></td>
		</tr>
		
		<%
			Dim Grnd_Ttl_ACT_THISMONTH, Grnd_Ttl_BUD_THISMONTH, Grnd_Ttl_ACT_TODATE, Grnd_Ttl_BUD_TODATE, Grnd_Ttl_BUD_ANNUAL
			Dim Grnd_Ttl_THISMONTHPER, Grnd_Ttl_TODATEPER
			Grnd_Ttl_ACT_THISMONTH		= Grnd_Ttl_ACT_THISMONTH + Ttl_ACT_THISMONTH
			Ttl_ACT_THISMONTH 			= 0
			Grnd_Ttl_BUD_THISMONTH		= Grnd_Ttl_BUD_THISMONTH + Ttl_BUD_THISMONTH
			Ttl_BUD_THISMONTH			= 0 
			Ttl_THISMONTHPER			= 0
			
			Grnd_Ttl_ACT_TODATE			= Grnd_Ttl_ACT_TODATE + Ttl_ACT_TODATE
			Ttl_ACT_TODATE				= 0
			Grnd_Ttl_BUD_TODATE			= Grnd_Ttl_BUD_TODATE + Ttl_BUD_TODATE
			Ttl_BUD_TODATE				= 0
			Ttl_TODATEPER				= 0
			
			Grnd_Ttl_BUD_ANNUAL			= Grnd_Ttl_BUD_ANNUAL + Ttl_BUD_ANNUAL
			Ttl_BUD_ANNUAL				= 0
			
			If CDbl(Grnd_Ttl_BUD_THISMONTH) <> 0 Then
				Grnd_Ttl_THISMONTHPER	= Grnd_Ttl_ACT_THISMONTH / Grnd_Ttl_BUD_THISMONTH * 100
			Else
				Grnd_Ttl_THISMONTHPER	= 0 
			END IF
			
			IF CDbl(Grnd_Ttl_BUD_TODATE) <> 0 Then
				Grnd_Ttl_TODATEPER			= Grnd_Ttl_ACT_TODATE / Grnd_Ttl_BUD_TODATE * 100
			Else
				Grnd_Ttl_TODATEPER			= 0 
			END IF
		Else
			NOX = NOX+ 1
		End IF
			
		rsMonthlyPupuk.MoveNext
		wend
	%>
	<tr>
			<td align="Center" style="background-color:#FFFF00;" colspan="<%=var_colspan%>"><b>Grand Total</b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_ACT_THISMONTH,2)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_BUD_THISMONTH,2)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_THISMONTHPER,2)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_ACT_TODATE,2)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_BUD_TODATE,2)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_TODATEPER,2)%></b></td>
			<td align="right" style="background-color:#FFFF00;"><b><%=FormatNumber(Grnd_Ttl_BUD_ANNUAL,2)%></b></td>
	</tr>
<%
	rsMonthlyPupuk.Close
	Set rsMonthlyPupuk = Nothing
	end if
%>
</tbody>
</table>
</body>
</html>