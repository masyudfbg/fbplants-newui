<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim txtDate, txtJob_forfieldstatus
	Dim strMain, nSpanTotal
	Dim rsAttDaily, strAttDaily,  v_txtJob_forfieldstatus
	Dim v_txtDate, NO, cClass
	Dim sprevFcba,SprevCompany,SprevRegion,SprevGroupRegion,SprevPlantation
	Dim tempfcba,tempCompany,tempRegion,tempGroupRegion,tempPlantation
	DIM q,FcbaName,CompanyName,MgrName,KTUName,CropName,FcbaType
	
	
	
	v_txtDate				=	request("txtDate")
	v_txtJob_forfieldstatus	=	request("txtJob_forfieldstatus")
	v_txtGroupPlantation   = Request("txtGroupPlantation")
	v_txtPlantation  	   = Request("txtPlantation")
	v_txtGroupRegion  	   = Request("txtGroupRegion")
	v_txtRegion  		   = Request("txtRegion")
	v_txtcompany  		   = Request("txtCompany")
	v_txtfcba 			   = Request("txtfcba")
	
	nSpanTotal = 6
	
   if v_txtDate ="" then
	   v_txtDate= textdate(Now())
    end if	

	 'response.write  "GP " & v_txtGroupPlantation & "  P " & v_txtPlantation & "  GR " & v_txtGroupRegion & " R " & v_txtRegion & " C " & v_txtcompany & " S " & v_txtfcba 
	 'response.write  "  StockGroup " &  v_txtStockGroup	 & "  Stockcode " & v_txtStock
	 
	        Dim selectFields,strFieldCons,selectCons,whereCons,orderCons 
	        strFieldCons = " USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',par_1.fcba) GroupPlantation,USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',par_1.fcba) Plantation, " & _
						   " USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',par_1.fcba) groupregion ,USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',par_1.fcba) region,USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',par_1.fcba) company, BUSINESSUNIT_API.get_OrderData(par_1.fcba) norut, par_1.fcba, " 

						   
			wherecons = " where company is not null  "
			
			if v_txtjob_forfieldstatus <> "--SUMMARY--" and v_txtjob_forfieldstatus <> "--ALL--"  and v_txtjob_forfieldstatus <> "" then
					whereCons = whereCons &  " AND job_forfieldstatus = '" & v_txtjob_forfieldstatus & "'"
			end if
			
			if v_txtGroupPlantation <> "--SUMMARY--"  and v_txtGroupPlantation <> "--ALL--" and v_txtGroupPlantation <> "" then
			   whereCons = wherecons & " And GroupPlantation = '" & v_txtGroupPlantation & "' " 
			end if
			if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
			    if v_txtPlantation <> "--ALL--" then
				   whereCons = whereCons & " And  Plantation = '" & v_txtPlantation & "' "
				end if   
				selectCons = selectCons & ",Plantation"
				orderCons  = orderCons & ", PLANTATION_API.get_OrderData(Plantation) "
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if v_txtGroupRegion <> "--SUMMARY--"   and v_txtGroupRegion <> "" then
			    if v_txtGroupRegion <> "--ALL--" then
				   whereCons = whereCons & " And  GroupRegion = '" & v_txtGroupRegion & "' "
				end if
				selectCons = selectCons & " ,GroupRegion"
				orderCons  = orderCons & ", GROUPREGION_API.get_OrderData(GroupRegion) "
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtRegion <> "--SUMMARY--"  and v_txtRegion <> "" then
			    if v_txtRegion <> "--ALL--" then
				   whereCons = whereCons & " And  Region = '" & v_txtRegion & "' "
				end if
				selectCons = selectCons & " ,Region"
				orderCons  = orderCons & ", REGION_API.get_OrderData(Region) "
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
			if  v_txtCompany <> "--SUMMARY--" and v_txtCompany <> "" then
			    if v_txtCompany  <> "--ALL--" then
				   whereCons = whereCons & " And  company = '" & v_txtCompany & "' "
				end if
				selectCons = selectCons & " ,company"
			    orderCons  = orderCons & ", COMPANY_API.get_OrderData(company) "
				q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where FCCOMPANYCODE = '" & v_txtCompany & "'"	
				set rsCat = connStr.Execute(q)
				if not rsCat.eof then
				   CompanyName=rsCat("FCCOMPANYNAME")			   
			    end if
			else
			    nSpanTotal = nSpanTotal - 1 	
			end if
	        if  v_txtfcba <> "--SUMMARY--"  and v_txtfcba <> "" then
			    if v_txtfcba  <> "--ALL--" then
			       whereCons = whereCons & " And  fcba = '" & v_txtfcba & "' "
				end if
			    selectCons = selectCons & ",fcba"
			    orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba) "
				q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
				set rsCat = connStr.Execute(q)
				if not rsCat.eof then
				   fcbaname = rsCat("fcName")
				   CompanyName=rsCat("FCCOMPANYNAME")
				   MgrName = rsCat("FCMANAGER")
				   KTUName = rsCat("FCKTU")
				   CropName=rsCat("FCCROP")
				   FcbaType=rsCat("fcType")
			    end if
			else
			    nSpanTotal = nSpanTotal - 1 
			end if
	        if request.Form("txtUKerja")="Tampil" then
			    selectCons = selectCons & ", division"
				orderCons  = orderCons & ",division "
			else
			    nSpanTotal = nSpanTotal - 1 
			end if   	
			   
	  ' strMain = " select " & strFieldCons & " param_1, job_forfieldstatus,decode(division,'',job_forfieldstatus,division) division, kbt_0, kht_0, khl_0, total_0," & _
				' " kbt_1, kht_1, khl_1, total_1,bud_shi," & _
				' " kbt_2, kht_2, khl_2, total_2,bud_sbi" & _
				' " from" & _
				' " (SELECT fcba, job_forfieldstatus,division, sum(kbt_0) kbt_0, sum(kht_0) kht_0, sum(khl_0) khl_0, sum(total_0) total_0," & _
				' " sum(kbt_1) kbt_1, sum(kht_1) kht_1, sum(khl_1) khl_1, sum(total_1) total_1,sum(bud_shi) bud_shi," & _
				' " sum(kbt_2) kbt_2, sum(kht_2) kht_2, sum(khl_2) khl_2, sum(total_2) total_2,sum(bud_sbi) bud_sbi" & _
				' " FROM (SELECT DISTINCT fcba, TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') fddate," & _
				' " 'VRA' job_forfieldstatus,'' division, 0 kbt_0, 0 kht_0, 0 khl_0," & _
				' " 0 total_0, 0 kbt_1, 0 kht_1, 0 khl_1, 0 total_1,0 bud_shi," & _
				' " 0 kbt_2, 0 kht_2, 0 khl_2, 0 total_2,0 bud_sbi" & _
				' " FROM job" & _
				' " UNION" & _
				' " SELECT DISTINCT fcba, TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') fddate," & _
				' " 'WRA' job_forfieldstatus,'' division, 0 kbt_0, 0 kht_0, 0 khl_0," & _
				' " 0 total_0, 0 kbt_1, 0 kht_1, 0 khl_1, 0 total_1,0 bud_shi," & _
				' " 0 kbt_2, 0 kht_2, 0 khl_2, 0 total_2,0 bud_sbi" & _
				' " FROM job" & _
				' " UNION" & _
				' " SELECT fcba, fddate, job_forfieldstatus,division, kbt kbt_0, kht kht_0," & _
				' " khl khl_0, (kbt + kht + khl) total_0, 0 kbt_1, 0 kht_1,0 bud_shi," & _
				' " 0 khl_1, 0 total_1, 0 kbt_2, 0 kht_2, 0 khl_2, 0 total_2,0 bud_sbi" & _
				' " FROM ATTENDANCE_DAILY_SUMMMARY_TAB" & _
				' " WHERE fddate = TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')" & _
				' " UNION" & _
				' " SELECT fcba, fddate, job_forfieldstatus,division, 0 kbt_0, 0 kht_0, 0 khl_0," & _
				' " 0 total_0, kbt kbt_1, kht kht_1, khl khl_1," & _
				' " (kbt + kht + khl) total_1,nvl(bud_hi,0) bud_shi, 0 kbt_2, 0 kht_2, 0 khl_2," & _
				' " 0 total_2,0 bud_sbi" & _
				' " FROM ATTENDANCE_DAILY_SUMMMARY_TAB" & _
				' " WHERE fddate >=" & _
				' " TO_DATE ( TO_CHAR (TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY'),'MM')" & _
				' " || '/01/'" & _
				' " || TO_CHAR (TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')," & _
				' " 'YYYY'" & _
				' " )," & _
				' " 'MM/DD/YYYY'" & _
				' " )" & _
				' " AND fddate <= TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') and to_char(fddate,'YYYY') = TO_CHAR (TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY'),'YYYY') " & _
				' " UNION" & _
				' " SELECT fcba, fddate, job_forfieldstatus,division, 0 kbt_0, 0 kht_0, 0 khl_0," & _
				' " 0 total_0, 0 kbt_1, 0 kht_1, 0 khl_1, 0 total_1,0 bud_shi, kbt kbt_2," & _
				' " kht kht_2, khl khl_2, (kbt + kht + khl) total_2,nvl(bud_hi,0) bud_sbi" & _ 
				' " FROM ATTENDANCE_DAILY_SUMMMARY_TAB" & _
				' " WHERE fddate <= TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') and to_char(fddate,'YYYY') = TO_CHAR (TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY'),'YYYY'))" & _
				' " GROUP BY fcba, job_forfieldstatus,division)" & _
				' " par_1," & _
				' " (select fcba, fccode, param_1 from parameter" & _
				' " where" & _
				' " param_1 is not null and" & _
				' " parcode = 'J0001'" & _
				' " order by param_1) par_2" & _
				' " where" & _
				' " par_1.job_forfieldstatus = par_2.fccode and" & _
				' " par_1.fcba = par_2.fcba" & _
				' " order by par_2.param_1" 

    'response.write strMain
	'strAttDaily	=	" SELECT  param_1 "
	' if len(trim(v_txtjob_forfieldstatus)) > 0 then 
		' if v_txtjob_forfieldstatus <> "--SUMMARY--" then
			' strAttDaily = strAttDaily & " , job_forfieldstatus,division "
		' end if
	' else
		' strAttDaily = strAttDaily & " , job_forfieldstatus,division  "
	' end if
	
	      strMain =  "  select fcba ,groupplantation,plantation,groupregion ,region ,company ,parameter_api.GetParam1(job_forfieldstatus,'J0001',fcba)  param_1,job_forfieldstatus,division,  " & _
					 "	sum(kbt_0) kbt_0, sum(kht_0) kht_0,sum(khl_0) khl_0,sum(total_0) total_0,  " & _
					 "	sum(kbt_1) kbt_1, sum(kht_1) kht_1,sum(khl_1) khl_1,sum(total_1) total_1,sum(bud_shi) bud_shi,  " & _
					 "	sum(kbt_2) kbt_2, sum(kht_2) kht_2,sum(khl_2) khl_2,sum(total_2) total_2,sum(bud_sbi) bud_sbi   " & _
					 "	from (  select fcba  ,groupplantation,plantation,groupregion ,region ,company , fddate, job_forfieldstatus,division, " & _
					 "			kbt kbt_0, kht kht_0, khl khl_0,(kbt+kht+khl) total_0, 0 kbt_1, 0 kht_1, 0 khl_1, 0 total_1,0 bud_shi, 0 kbt_2, " & _
					 "			0 kht_2, 0 khl_2, 0 total_2,0 bud_sbi " & _
					 "			from ATTENDANCE_DAILY_SUMMMARY_TAB where fddate = TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') " & _
					 "			union " & _
					 "			select fcba  ,groupplantation,plantation,groupregion ,region ,company , fddate, job_forfieldstatus,division, " & _
					 "			 0 kbt_0, 0 kht_0, 0 khl_0,0 total_0, kbt kbt_1,kht  kht_1, khl khl_1,(kbt + kht + khl) total_1,bud_hi bud_shi, 0 kbt_2, " & _
					 "			0 kht_2, 0 khl_2, 0 total_2,0 bud_sbi  " & _
					 "			from ATTENDANCE_DAILY_SUMMMARY_TAB where fddate <= TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY')  " & _
					 "			and to_char(fddate,'YYYY') = TO_CHAR (TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY'),'YYYY') " & _
					 "			and to_char(fddate,'MM')  = TO_CHAR (TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY'),'MM') " & _
					 "			union " & _
					 "			select fcba  ,groupplantation,plantation,groupregion ,region ,company , fddate, job_forfieldstatus,division, " & _
					 "			0 kbt_0, 0 kht_0, 0 khl_0,0 total_0, 0 kbt_1, 0 kht_1, 0 khl_1, 0 total_1,0 bud_shi, kbt kbt_2, " & _
					 "			kht kht_2, khl khl_2, kbt + kht + khl total_2,nvl(bud_hi,0) bud_sbi " & _
					 "			from ATTENDANCE_DAILY_SUMMMARY_TAB where fddate <= TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY') " & _
					 "			and to_char(fddate,'YYYY') = TO_CHAR (TO_DATE ('" & v_txtDate & "', 'MM/DD/YYYY'),'YYYY') " & _
					 "	) group by fcba  ,groupplantation,plantation,groupregion ,region ,company , job_forfieldstatus,division " & _
					 " 	order by parameter_api.GetParam1(job_forfieldstatus,'J0001',fcba)  "
	
	
	strAttDaily	=	" , sum(kbt_0) kbt_0, sum(kht_0) kht_0, sum(khl_0) khl_0, sum(total_0) total_0, " & _
					" sum(kbt_1) kbt_1, sum(kht_1) kht_1, sum(khl_1) khl_1, sum(total_1) total_1, sum(nvl(bud_shi,0)) bud_shi,sum(nvl(bud_sbi,0)) bud_sbi,  " & _
					" sum(kbt_2) kbt_2, sum(kht_2) kht_2, sum(khl_2) khl_2, sum(total_2) total_2 "
					
	strAttDaily = "SELECT param_1,job_forfieldstatus,GroupPlantation " & selectCons & strAttDaily	& " FROM ( " & strMain & " )  " & whereCons & "  group by  param_1,job_forfieldstatus,GroupPlantation" & selectCons & "  Order by param_1,job_forfieldstatus,GroupPlantation" & orderCons  
	
	    'strAttDaily = strAttDaily & " WHERE fcba = '" & Session("fcba") & "' "
	
	' if len(trim(v_txtjob_forfieldstatus)) > 0 then 
		' if v_txtjob_forfieldstatus <> "--SUMMARY--" then
			' strAttDaily = strAttDaily & " AND job_forfieldstatus = '" & v_txtjob_forfieldstatus & "'"
		' end if
	' end if
	
	' strAttDaily = strAttDaily & " GROUP BY param_1 "
	' if len(trim(v_txtjob_forfieldstatus)) > 0 then 
		' if v_txtjob_forfieldstatus <> "--SUMMARY--" then
			' strAttDaily = strAttDaily & " , job_forfieldstatus ,division"
		' end if
	' else
		' strAttDaily = strAttDaily & " , job_forfieldstatus,division "
	' end if
	
	' strAttDaily = strAttDaily & " ORDER BY  param_1 "
	' if len(trim(v_txtjob_forfieldstatus)) > 0 then 
		' if v_txtjob_forfieldstatus <> "--SUMMARY--" then
			' strAttDaily = strAttDaily & " , job_forfieldstatus ,division"
		' end if
	' else
		' strAttDaily = strAttDaily & " , job_forfieldstatus,division "
	' end if

   'response.write strAttDaily
	Set rsAttDaily = Server.CreateObject("ADODB.Recordset")
	rsAttDaily.Open strAttDaily, connstr, adOpenStatic 
	
%>


<html>

<head>

<title>ePlaS Reporting - Laporan Harian Tenaga Kerja</title>
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
</head>

<body>
<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
<form name="frmEntry" method="post" action="p_AttendanceDaily.asp">
	<div id="wrapper">
	<div id="header-wrapper">
	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/Dashboard.asp">Back to Report</a></td>
		</tr>
	</table>
	
	
	</div>
<!--#include file="FormSelectNewSum.html"-->		
<table id="form-table" class="centered-margin">
	<tr>
		<th>TANGGAL</th>
		<th>URAIAN</th>
		<th>UNIT KERJA</th>
		<th>&nbsp;</th>
	</tr>
	<tr>		
	
		<td align="center">
		<% if v_txtDate <> "" and v_txtDate = v_txtDate then %>
			<input class="datepicker" type="text" maxlength="10" name="txtDate" size="10" value="<%=v_txtDate%>" title="Type Date mm/dd/yyyy" >
		<% else %>
			<input class="datepicker" type="text" maxlength="10" name="txtDate" size="10" value="<%=textDate(Date)%>" title="Type Date mm/dd/yyyy" >
		<% end if%>
			<!--<a href="javascript:show_calendar('frmEntry.txtDate');" onMouseOver="window.status='Date Picker';return true;" onMouseOut="window.status='';return true;">
				<img src="../Images/show-calendar.gif" width=24 height=22 border=0 align=top>
			</a>-->
		</td>

		<td align="center">
		<select name="txtJob_forfieldstatus">	
		<option value="">--ALL--</option>
		<%
			Dim rsJob_forfieldstatus
			Set rsJob_forfieldstatus = Server.CreateObject("ADODB.Recordset")
			rsJob_forfieldstatus.CursorLocation = adUseClient
			
		    if len(trim(v_txtFcba)) > 0 then 
				strsql =  " SELECT 'VRA' as Job_forfieldstatus from DUAL UNION SELECT 'WRA' as Job_forfieldstatus from DUAL UNION SELECT DISTINCT DECODE(Job_forfieldstatus,'Balance Sheet','LAIN-LAIN',Job_forfieldstatus) as Job_forfieldstatus FROM JOB WHERE user_company_api.IsAuthorizedFcba('" & Session("CodeName") & "',fcba) is not null "
			else
				strsql =  " SELECT 'VRA' as Job_forfieldstatus from DUAL UNION SELECT 'WRA' as Job_forfieldstatus from DUAL UNION SELECT DISTINCT DECODE(Job_forfieldstatus,'Balance Sheet','LAIN-LAIN',Job_forfieldstatus) as Job_forfieldstatus FROM JOB WHERE user_company_api.IsAuthorizedFcba('" & Session("CodeName") & "',fcba) is not null "
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
		
		
		
		<td>
	
			<select name="txtUKerja">
			<option value="Tutup" <%if request.Form("txtUKerja")="Tutup" Then%> Selected <%End if%>>Tidak Tampilkan</option>
			<option value="Tampil" <%if request.Form("txtUKerja")="Tampil" Then%> Selected <%End if%>>Tampilkan</option>
			</select>
		</td>
		
		<td align="center"><input type="submit" value="View" /></td>
	</tr>
</table>
<div id="wrapper">
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
<tr>
<th colspan="21" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
<div id="header-wrapper">
	<div id="header">
			 <% if  CompanyName <>  "" then %>
					<h4>PT.<%=CompanyName%></h4> 
			 <% End if%>
			 <% If  FcbaName <> ""  then %>
					   <h5><%=FcbaName%></h5>   
			 <% End if%>
	</div>
	<div id="subheader"class="fixed-1 clearfix centered">
			<div id="subheader" class="centered">
				<h2>LAPORAN HARIAN TENAGA KERJA</h3>
				<h5>TANGGAL: <%=v_txtDate%></h5>
			</div>
	</div>
</div>
</th>
</tr>
	<tr>
		
		<% if trim(request("txtjob_forfieldstatus")) <> "--SUMMARY--" then %>
		    <th rowspan="2" style="background-color:#00A300;color:#fff;">URAIAN</th>
		<% else
			nSpanTotal = nSpanTotal - 1 			
	    end if %>
		
		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
		<% end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
		<% end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
		<% end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
		<% end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
		<% end if %>
		<% if request.Form("txtUKerja")="Tampil" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;">UNIT KERJA</th>
		<% end if %>
		<% if v_txtDate <> "" and v_txtDate = v_txtDate then %>
		<th colspan="4" style="background-color:#00A300;color:#fff;">HARI INI <%=v_txtDate%></th>
		<% else %>
		<th colspan="4" style="background-color:#00A300;color:#fff;">HARI INI</th>
		<% end if%>
		<% if v_txtDate <> "" and v_txtDate = v_txtDate then %>
		<th colspan="4" style="background-color:#00A300;color:#fff;">SAMPAI HARI INI <%=v_txtDate%></th>
		<% else %>
		<th colspan="4" style="background-color:#00A300;color:#fff;">SAMPAI HARI INI</th>
		<% end if%>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;">SAMPAI DENGAN BULAN INI</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">BUDGET</th>
		
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;">KBT</th>
		<th style="background-color:#00A300;color:#fff;">KHT</th>
		<th style="background-color:#00A300;color:#fff;">KHL</th>
		<th style="background-color:#00A300;color:#fff;">TOTAL</th>
		<th style="background-color:#00A300;color:#fff;">KBT</th>
		<th style="background-color:#00A300;color:#fff;">KHT</th>
		<th style="background-color:#00A300;color:#fff;">KHL</th>
		<th style="background-color:#00A300;color:#fff;">TOTAL</th>
		<th style="background-color:#00A300;color:#fff;">KBT</th>
		<th style="background-color:#00A300;color:#fff;">KHT</th>
		<th style="background-color:#00A300;color:#fff;">KHL</th>
		<th style="background-color:#00A300;color:#fff;">TOTAL</th>
	</tr>
        </thead>
 <tbody>
<%
Dim nkbt_0, nkht_0, nkhl_0, nTotal_0
Dim nkbt_1, nkht_1, nkhl_1, nTotal_1,nsbi,nshi
Dim nkbt_2, nkht_2, nkhl_2, nTotal_2
Dim sprevjob_forfieldstatus, tempjob_forfieldstatus
	if not rsAttDaily.EOF then
		NO = 1
		nkbt_0 = 0
		nkht_0 = 0
		nkhl_0 = 0
		nTotal_0 = 0
		nsbi = 0
		nshi = 0
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
					
				
				<% if v_txtjob_forfieldstatus <> "--SUMMARY--" then %>
				    <%     tempjob_forfieldstatus = rsAttDaily("job_forfieldstatus")			   
						   tempPlantation = "--ALL--"
						   tempGroupRegion = "--SUMMARY--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
					 %>
							
				    <% if sprevjob_forfieldstatus <> rsAttDaily("job_forfieldstatus") then %>
				        <td align="left">
						    <a href="p_AttendanceDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtJob_forfieldstatus=<%=tempjob_forfieldstatus%>&txtdate=<%=v_txtdate%>">
						     <%=rsAttDaily("job_forfieldstatus")%>
							</a> 
					    </td>
						<% sprevjob_forfieldstatus = rsAttDaily("job_forfieldstatus") %>
					<% else	%>
					    <td align="left"></td>					
					<% end if %>							 
				<% end if %>
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<%     tempjob_forfieldstatus = rsAttDaily("job_forfieldstatus")			   
						   tempPlantation = rsAttDaily("Plantation")
						   tempGroupRegion = "--ALL--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
					 %>
					<% if SprevPlantation <> rsAttDaily("Plantation") then %>
						<td align="center">
							<a href="p_AttendanceDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtJob_forfieldstatus=<%=tempjob_forfieldstatus%>&txtdate=<%=v_txtdate%>">
							<%=rsAttDaily("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = rsAttDaily("plantation") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
				    <%  tempjob_forfieldstatus = rsAttDaily("job_forfieldstatus")
					    if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
							tempPlantation = rsAttDaily("Plantation")
						else
							tempPlantation = ""
						end if
					   tempGroupRegion = rsAttDaily("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					  
					%>
					
					<% if SprevGroupRegion <> rsAttDaily("GroupRegion") then %>
						<td align="center">
							<a href="p_AttendanceDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtJob_forfieldstatus=<%=tempjob_forfieldstatus%>&txtdate=<%=v_txtdate%>">
							<%=rsAttDaily("groupregion")%>
							</a>
						</td>
						<% SprevGroupRegion = rsAttDaily("GroupRegion") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <%  tempjob_forfieldstatus = rsAttDaily("job_forfieldstatus")
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsAttDaily("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsAttDaily("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsAttDaily("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   
					%>
					
					<% if SprevRegion <> rsAttDaily("Region") then %>
						<td align="center">
							<a href="p_AttendanceDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtJob_forfieldstatus=<%=tempjob_forfieldstatus%>&txtdate=<%=v_txtdate%>">
							<%=rsAttDaily("region")%>
							</a>
						</td>
						<% SprevRegion = rsAttDaily("region") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% 
					    tempjob_forfieldstatus = rsAttDaily("job_forfieldstatus")
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsAttDaily("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsAttDaily("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsAttDaily("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rsAttDaily("Company")
					   tempfcba = "--ALL--"
					   				   
					 %>
					
					<% if SprevCompany <> rsAttDaily("Company") then %>
						<td align="center">
							<a href="p_AttendanceDaily.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtJob_forfieldstatus=<%=tempjob_forfieldstatus%>&txtdate=<%=v_txtdate%>">
							<%=rsAttDaily("company")%>
							</a>
						</td>
						<% SprevCompany = rsAttDaily("Company") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <%  tempjob_forfieldstatus = rsAttDaily("job_forfieldstatus")
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  ""  then
					      tempPlantation = rsAttDaily("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   if  v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  ""  then
					       tempGroupRegion = rsAttDaily("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
					       tempRegion = rsAttDaily("Region") 
					   else
					       tempRegion = ""
					   end if
					   if  v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then
					       tempCompany =rsAttDaily("Company")
					   else
					       tempCompany =""
					   end if
					   tempfcba = rsAttDaily("fcba")
					   		   
					 %>
					
					<% if sprevFcba <> rsAttDaily("fcba") then %>
						<td align="center">
							  <%=rsAttDaily("fcba")%>
						</td>
						<% sprevFcba = rsAttDaily("fcba") %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	

				<% end if %>
				<% if request.Form("txtUKerja")="Tampil" then %>
				<td align="left"><%=rsAttDaily("division")%></td>
				<% end if %>
				<td align="right"><%=FormatNumber(rsAttDaily("kbt_0"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("kht_0"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("khl_0"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("total_0"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsAttDaily("kbt_1"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("kht_1"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("khl_1"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("total_1"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("bud_shi"),2)%></td>	
				<td align="right"><%=FormatNumber(rsAttDaily("kbt_2"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("kht_2"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("khl_2"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("total_2"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("bud_sbi"),2)%></td>	
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
            nshi = nshi + + Cdbl(rsAttDaily("bud_shi"))            
			
			nkbt_2 = nkbt_2 + Cdbl(rsAttDaily("kbt_2"))
			nkht_2 = nkht_2 + Cdbl(rsAttDaily("kht_2"))
			nkhl_2 = nkhl_2 + Cdbl(rsAttDaily("khl_2"))
			nTotal_2 = nTotal_2 + Cdbl(rsAttDaily("total_2"))
			 nsbi = nsbi + + Cdbl(rsAttDaily("bud_sbi"))  
			
			rsAttDaily.MoveNext
		Wend
	end if
	rsAttDaily.Close
	set rsAttDaily = Nothing
%>
			<tr bgcolor="#F4F4A0">
				<td align="center" colspan="<%= 1 + nSpanTotal %>">TOTAL</td>
				<td align="right"><%=FormatNumber(nkbt_0,2)%></td>
				<td align="right"><%=FormatNumber(nkht_0,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_0,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_0,2)%></td>

				<td align="right"><%=FormatNumber(nkbt_1,2)%></td>
				<td align="right"><%=FormatNumber(nkht_1,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_1,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_1,2)%></td>
                <td align="right"><%=FormatNumber(nshi,2)%> </td>	 
				<td align="right"><%=FormatNumber(nkbt_2,2)%></td>
				<td align="right"><%=FormatNumber(nkht_2,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_2,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_2,2)%></td>
				 <td align="right"><%=FormatNumber(nsbi,2)%> </td>	
				
			</tr>
<tbody>
</table>
</form>

</div>
</body>
</html>