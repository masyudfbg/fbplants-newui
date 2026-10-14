<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
	Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i,v_txtyear, v_txtmonth,txtDate, txtJob_forfieldstatus
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
	
	        v_txtyear 		= Request("txtyear")
			v_txtmonth 		= lpad(Request("txtmonth"),"0",2)

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
			   
	  strMain = " select year,groupplantation,plantation,groupregion,region,company,t.fcba,h.param_1,job_forfieldstatus,division,locationcode, " & _
				"		   decode(month,'01',sum(nvl(kbt,0)),0) kbt_01,decode(month,'01',sum(nvl(kht,0)),0) kht_01,decode(month,'01',sum(nvl(khl,0)),0) khl_01,decode(month,'01',sum(nvl(bud_month,0)),0) bud_month_01, " & _
				"		   decode(month,'02',sum(nvl(kbt,0)),0) kbt_02,decode(month,'02',sum(nvl(kht,0)),0) kht_02,decode(month,'02',sum(nvl(khl,0)),0) khl_02,decode(month,'02',sum(nvl(bud_month,0)),0) bud_month_02, " & _
				"		   decode(month,'03',sum(nvl(kbt,0)),0) kbt_03,decode(month,'03',sum(nvl(kht,0)),0) kht_03,decode(month,'03',sum(nvl(khl,0)),0) khl_03,decode(month,'03',sum(nvl(bud_month,0)),0) bud_month_03, " & _
				"		   decode(month,'04',sum(nvl(kbt,0)),0) kbt_04,decode(month,'04',sum(nvl(kht,0)),0) kht_04,decode(month,'04',sum(nvl(khl,0)),0) khl_04,decode(month,'04',sum(nvl(bud_month,0)),0) bud_month_04, " & _
				"		   decode(month,'05',sum(nvl(kbt,0)),0) kbt_05,decode(month,'05',sum(nvl(kht,0)),0) kht_05,decode(month,'05',sum(nvl(khl,0)),0) khl_05,decode(month,'05',sum(nvl(bud_month,0)),0) bud_month_05, " & _
				"		   decode(month,'06',sum(nvl(kbt,0)),0) kbt_06,decode(month,'06',sum(nvl(kht,0)),0) kht_06,decode(month,'06',sum(nvl(khl,0)),0) khl_06,decode(month,'06',sum(nvl(bud_month,0)),0) bud_month_06, " & _
				"		   decode(month,'07',sum(nvl(kbt,0)),0) kbt_07,decode(month,'07',sum(nvl(kht,0)),0) kht_07,decode(month,'07',sum(nvl(khl,0)),0) khl_07,decode(month,'07',sum(nvl(bud_month,0)),0) bud_month_07, " & _
				"		   decode(month,'08',sum(nvl(kbt,0)),0) kbt_08,decode(month,'08',sum(nvl(kht,0)),0) kht_08,decode(month,'08',sum(nvl(khl,0)),0) khl_08,decode(month,'08',sum(nvl(bud_month,0)),0) bud_month_08, " & _
				"		   decode(month,'09',sum(nvl(kbt,0)),0) kbt_09,decode(month,'09',sum(nvl(kht,0)),0) kht_09,decode(month,'09',sum(nvl(khl,0)),0) khl_09,decode(month,'09',sum(nvl(bud_month,0)),0) bud_month_09, " & _
				"		   decode(month,'10',sum(nvl(kbt,0)),0) kbt_10,decode(month,'10',sum(nvl(kht,0)),0) kht_10,decode(month,'10',sum(nvl(khl,0)),0) khl_10,decode(month,'10',sum(nvl(bud_month,0)),0) bud_month_10, " & _
				"		   decode(month,'11',sum(nvl(kbt,0)),0) kbt_11,decode(month,'11',sum(nvl(kht,0)),0) kht_11,decode(month,'11',sum(nvl(khl,0)),0) khl_11,decode(month,'11',sum(nvl(bud_month,0)),0) bud_month_11, " & _
				"		   decode(month,'12',sum(nvl(kbt,0)),0) kbt_12,decode(month,'12',sum(nvl(kht,0)),0) kht_12,decode(month,'12',sum(nvl(khl,0)),0) khl_12,decode(month,'12',sum(nvl(bud_month,0)),0) bud_month_12 " & _
				"	from ATTEND_MONTHLY_SUMMARY_TAB   t,  " & _
				"		( select fcba, fccode, param_1  " & _
				"		  from parameter  " & _
				"		  where param_1 is not null and parcode = 'J0001' ) h  " & _
				"	where t.year ='" & v_txtyear &"' and t.month <= '"& v_txtmonth &"'  " & _
				"		  and t.job_forfieldstatus = h.fccode and t.fcba = h.fcba  " & _
				"	group by year,month,groupplantation,plantation,groupregion,region,company,t.fcba,h.param_1,job_forfieldstatus,division,locationcode" 

   
	strAttDaily	=	" , sum(kbt_01) kbt_01, sum(kht_01) kht_01, sum(khl_01) khl_01,( sum(kbt_01) + sum(kht_01) + sum(khl_01) ) total_01 " & _
					" , sum(kbt_02) kbt_02, sum(kht_02) kht_02, sum(khl_02) khl_02,( sum(kbt_02) + sum(kht_02) + sum(khl_02) ) total_02 "& _
					" , sum(kbt_03) kbt_03, sum(kht_03) kht_03, sum(khl_03) khl_03,( sum(kbt_03) + sum(kht_03) + sum(khl_03) ) total_03 "& _
					" , sum(kbt_04) kbt_04, sum(kht_04) kht_04, sum(khl_04) khl_04,( sum(kbt_04) + sum(kht_04) + sum(khl_04) ) total_04 "& _
					" , sum(kbt_05) kbt_05, sum(kht_05) kht_05, sum(khl_05) khl_05,( sum(kbt_05) + sum(kht_05) + sum(khl_05) ) total_05 "& _
					" , sum(kbt_06) kbt_06, sum(kht_06) kht_06, sum(khl_06) khl_06,( sum(kbt_06) + sum(kht_06) + sum(khl_06) ) total_06 "& _
					" , sum(kbt_07) kbt_07, sum(kht_07) kht_07, sum(khl_07) khl_07,( sum(kbt_07) + sum(kht_07) + sum(khl_07) ) total_07 "& _
					" , sum(kbt_08) kbt_08, sum(kht_08) kht_08, sum(khl_08) khl_08,( sum(kbt_08) + sum(kht_08) + sum(khl_08) ) total_08 "& _
					" , sum(kbt_09) kbt_09, sum(kht_09) kht_09, sum(khl_09) khl_09,( sum(kbt_09) + sum(kht_09) + sum(khl_09) ) total_09 "& _
					" , sum(kbt_10) kbt_10, sum(kht_10) kht_10, sum(khl_10) khl_10,( sum(kbt_10) + sum(kht_10) + sum(khl_10) ) total_10 "& _
					" , sum(kbt_11) kbt_11, sum(kht_11) kht_11, sum(khl_11) khl_11,( sum(kbt_11) + sum(kht_11) + sum(khl_11) ) total_11 "& _
					" , sum(kbt_12) kbt_12, sum(kht_12) kht_12, sum(khl_12) khl_12,( sum(kbt_12) + sum(kht_12) + sum(khl_12) ) total_12 "& _
					" , ( sum(kbt_01) + sum(kbt_02) + sum(kbt_03) + sum(kbt_04) + sum(kbt_05) + sum(kbt_06) + sum(kbt_07) + sum(kbt_08) + sum(kbt_09) + sum(kbt_10) + sum(kbt_11) + sum(kbt_12)) kbt_SBI "& _
					" , ( sum(kht_01) + sum(kht_02) + sum(kht_03) + sum(kht_04) + sum(kht_05) + sum(kht_06) + sum(kht_07) + sum(kht_08) + sum(kht_09) + sum(kht_10) + sum(kht_11) + sum(kht_12)) kht_SBI "& _
					" , ( sum(khl_01) + sum(khl_02) + sum(khl_03) + sum(khl_04) + sum(khl_05) + sum(khl_06) + sum(khl_07) + sum(khl_08) + sum(khl_09) + sum(khl_10) + sum(khl_11) + sum(khl_12)) khl_SBI "& _
					" , ( sum(bud_month_01) + sum(bud_month_02) + sum(bud_month_03) +sum(bud_month_04) + sum(bud_month_05) + sum(bud_month_06) + sum(bud_month_10) + sum(bud_month_08) + sum(bud_month_09) + sum(bud_month_10) + sum(bud_month_11) + sum(bud_month_12) )  bud_SBI "
					
	strAttDaily = "SELECT param_1,job_forfieldstatus,GroupPlantation " & selectCons & strAttDaily	& " FROM ( " & strMain & " )  " & whereCons & "  group by  param_1,job_forfieldstatus,GroupPlantation" & selectCons & "  Order by param_1,job_forfieldstatus,GroupPlantation" & orderCons  
	
	 

   ' response.write strAttDaily
	Set rsAttDaily = Server.CreateObject("ADODB.Recordset")
	rsAttDaily.Open strAttDaily, connstr, adOpenStatic 
	
%>


<html>

<head>

<title>ePlaS Reporting - Laporan Harian Tenaga Kerja</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<!--<script src="/js/jquery-1.9.1.js"></script>
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

<!--<script type="text/javascript" src="http://jqueryjs.googlecode.com/files/jquery-1.3.1.min.js" > </script> -->
<script type="text/javascript">

    function PrintElem(elem)
    {
        Popup($(elem).html());
    }

    function Popup(data) 
    {
        var mywindow = window.open('', 'my div', 'height=400,width=600');
        mywindow.document.write('<html><head><title>my div</title>');
        mywindow.document.write('<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />');
        mywindow.document.write('</head><body >');
        mywindow.document.write(data);
        mywindow.document.write('</body></html>');

        mywindow.print();
        mywindow.close();

        return true;
    }

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
<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />
 <script src="http://code.jquery.com/jquery-1.9.1.js"></script>
<script src="http://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>
<script type="text/javascript">
$( "#myProgressBar" ).progressbar({
create: function( event, ui ) {}
});
</script>
</head>

<body>
<div class="top_select" style=" position:fixed; z-index:-999;  top:0; left:0px; width:100%;"> 
<div id="progressbar"></div>
<form name="frmEntry" method="post" action="ZZZp_Attendancemonthly.asp">

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
		<th>&nbsp;</th>
		<th>BULAN</th>
		<th>TAHUN</th>
		<th>URAIAN</th>
		<th>UNIT KERJA</th>
		<th>&nbsp;</th>
	</tr>
	<tr>		
	
		<td align="center">
		<% 
					MM=Request("txtMonth")
					yyyy=Request("txtYear")

					mon = Month(Session("v_Costbook"))
					If len(MM)=0 then MM=mon

					Yr = Year(Session("v_Costbook"))
					MinYr = Yr - 2
					MaxYr = Yr  + 2
					If len(yyyy)= 0 then yyyy = Yr
				%>
		
		   <td>
				<font size="2" color="#0000FF" ></font>
					<select name="txtMonth">
						<%
							For i = 1 to 12
						%>
							<Option Value="<%=i%>" <%If cint(i) = cint(MM) then%> Selected <%End If%>><%=MonthName(i)%></option>
						<%
							Next
						%>
					</select>
		  </td>	
		   <td>
				<font size="2"  color="#0000FF" ></font>
				
				<select name="txtYear">
					<%
						For i = MinYr to MaxYr
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(yyyy) then%> Selected <%End If%>><%=i%></option>
					<%
						Next
					%>
				</select>
			</td>

		<td align="center">
		<select name="txtJob_forfieldstatus">	
		<option value="">--ALL--</option>
		<%
			Dim rsJob_forfieldstatus
			Set rsJob_forfieldstatus = Server.CreateObject("ADODB.Recordset")
			rsJob_forfieldstatus.CursorLocation = adUseClient
			
		    if len(trim(v_txtFcba)) > 0 then 
				strsql =  " SELECT '--SUMMARY--' as Job_forfieldstatus from DUAL UNION SELECT 'VRA' as Job_forfieldstatus from DUAL UNION SELECT 'WRA' as Job_forfieldstatus from DUAL UNION SELECT DISTINCT DECODE(Job_forfieldstatus,'Balance Sheet','LAIN-LAIN',Job_forfieldstatus) as Job_forfieldstatus FROM JOB WHERE user_company_api.IsAuthorizedFcba('" & Session("CodeName") & "',fcba) is not null "
			else
				strsql =  " SELECT '--SUMMARY--' as Job_forfieldstatus from DUAL UNION SELECT 'VRA' as Job_forfieldstatus from DUAL UNION SELECT 'WRA' as Job_forfieldstatus from DUAL UNION SELECT DISTINCT DECODE(Job_forfieldstatus,'Balance Sheet','LAIN-LAIN',Job_forfieldstatus) as Job_forfieldstatus FROM JOB WHERE user_company_api.IsAuthorizedFcba('" & Session("CodeName") & "',fcba) is not null "
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
    <tr>
<td colspan="140" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
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
				<h2>LAPORAN BULANAN TENAGA KERJA</h3>
				<h5>PERIODE : <%=v_txtmonth%> - <%=v_txtyear%> </h5>
			</div>
	</div>
</div>
</td>
</tr>
</table>
</div>

<div id="mydiv" style="margin-top:280px;">
<table border="0" width="100%" id="product-table" class="xthf-site" style="background-color:#FFFFFF;">
<thead>
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
		
		<th colspan="4" style="background-color:#00A300;color:#fff;" >JAN</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >FEB</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >MAR</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >APR</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >MAY</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >JUN</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >JUL</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >AUG</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >SEP</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >OCT</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >NOV</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >DES</th>
		<th colspan="4" style="background-color:#00A300;color:#fff;" >SBI</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">BGT</th>
		
	</tr>
	<tr>
	    <% DIM K
		   FOR K = 1 TO 13 %>
			<th style="background-color:#00A300;color:#fff;">KBT</th>
			<th style="background-color:#00A300;color:#fff;">KHT</th>
			<th style="background-color:#00A300;color:#fff;">KHL</th>
			<th style="background-color:#00A300;color:#fff;">TOTAL</th>
		<% NEXT %>	
		
		
	</tr>
        </thead>
 <tbody>
<%
Dim nkbt_01, nkht_01, nkhl_01, nTotal_01,nkbt_02, nkht_02, nkhl_02, nTotal_02,nkbt_03, nkht_03, nkhl_03, nTotal_03,nkbt_04, nkht_04, nkhl_04, nTotal_04
Dim nkbt_05, nkht_05, nkhl_05, nTotal_05,nkbt_06, nkht_06, nkhl_06, nTotal_06,nkbt_07, nkht_07, nkhl_07, nTotal_07,nkbt_08, nkht_08, nkhl_08, nTotal_08
Dim nkbt_09, nkht_09, nkhl_09, nTotal_09,nkbt_10, nkht_10, nkhl_10, nTotal_10,nkbt_11, nkht_11, nkhl_11, nTotal_11,nkbt_12, nkht_12, nkhl_12, nTotal_12
Dim nkbt_SBI, nkht_SBI, nkhl_SBI, nTotal_SBI,nTotal_BUD
Dim Total_SBI
Dim nkbt_1, nkht_1, nkhl_1, nTotal_1,nsbi,nshi
Dim nkbt_2, nkht_2, nkhl_2, nTotal_2
Dim sprevjob_forfieldstatus, tempjob_forfieldstatus
	if not rsAttDaily.EOF then
		NO = 1
		
		
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
						    <a href="p_Attendancemonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtJob_forfieldstatus=<%=tempjob_forfieldstatus%>&txtdate=<%=v_txtdate%>">
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
							<a href="p_Attendancemonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtJob_forfieldstatus=<%=tempjob_forfieldstatus%>&txtdate=<%=v_txtdate%>">
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
							<a href="p_Attendancemonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtJob_forfieldstatus=<%=tempjob_forfieldstatus%>&txtdate=<%=v_txtdate%>">
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
							<a href="p_Attendancemonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtJob_forfieldstatus=<%=tempjob_forfieldstatus%>&txtdate=<%=v_txtdate%>">
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
							<a href="p_Attendancemonthly.asp?txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtJob_forfieldstatus=<%=tempjob_forfieldstatus%>&txtdate=<%=v_txtdate%>">
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
				
				<td align="right"><%=FormatNumber(rsAttDaily("kbt_01"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("kht_01"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("khl_01"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("total_01"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsAttDaily("kbt_02"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("kht_02"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("khl_02"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("total_02"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsAttDaily("kbt_03"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("kht_03"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("khl_03"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("total_03"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsAttDaily("kbt_04"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("kht_04"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("khl_04"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("total_04"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsAttDaily("kbt_05"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("kht_05"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("khl_05"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("total_05"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsAttDaily("kbt_06"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("kht_06"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("khl_06"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("total_06"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsAttDaily("kbt_07"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("kht_07"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("khl_07"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("total_07"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsAttDaily("kbt_08"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("kht_08"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("khl_08"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("total_08"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsAttDaily("kbt_09"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("kht_09"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("khl_09"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("total_09"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsAttDaily("kbt_10"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("kht_10"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("khl_10"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("total_10"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsAttDaily("kbt_11"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("kht_11"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("khl_11"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("total_11"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsAttDaily("kbt_12"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("kht_12"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("khl_12"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("total_12"),2)%></td>
				
				<td align="right"><%=FormatNumber(rsAttDaily("kbt_SBI"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("kht_SBI"),2)%></td>
				<td align="right"><%=FormatNumber(rsAttDaily("khl_SBI"),2)%></td>
				<% total_SBI = cdbl(rsAttDaily("kbt_SBI")) + cdbl(rsAttDaily("kht_SBI")) + cdbl(rsAttDaily("khl_SBI")) %>
				<td align="right"><%=FormatNumber(total_SBI,2)%></td>
				
				<td align="right"><%=FormatNumber(rsAttDaily("bud_SBI"),2)%></td>
				
			</tr>
<%
			NO = NO + 1			
			nkbt_01 = nkbt_01 + Cdbl(rsAttDaily("kbt_01"))
			nkht_01 = nkht_01 + Cdbl(rsAttDaily("kht_01"))
			nkhl_01 = nkhl_01 + Cdbl(rsAttDaily("khl_01"))
			nTotal_01 = nTotal_01 + Cdbl(rsAttDaily("total_01"))
			
			nkbt_02 = nkbt_02 + Cdbl(rsAttDaily("kbt_02"))
			nkht_02 = nkht_02 + Cdbl(rsAttDaily("kht_02"))
			nkhl_02 = nkhl_02 + Cdbl(rsAttDaily("khl_02"))
			nTotal_02 = nTotal_02 + Cdbl(rsAttDaily("total_02"))
			
			nkbt_03 = nkbt_03 + Cdbl(rsAttDaily("kbt_03"))
			nkht_03 = nkht_03 + Cdbl(rsAttDaily("kht_03"))
			nkhl_03 = nkhl_03 + Cdbl(rsAttDaily("khl_03"))
			nTotal_03 = nTotal_03 + Cdbl(rsAttDaily("total_03"))
			
			nkbt_04 = nkbt_04 + Cdbl(rsAttDaily("kbt_04"))
			nkht_04 = nkht_04 + Cdbl(rsAttDaily("kht_04"))
			nkhl_04 = nkhl_04 + Cdbl(rsAttDaily("khl_04"))
			nTotal_04 = nTotal_04 + Cdbl(rsAttDaily("total_04"))
			
			nkbt_05 = nkbt_05 + Cdbl(rsAttDaily("kbt_05"))
			nkht_05 = nkht_05 + Cdbl(rsAttDaily("kht_05"))
			nkhl_05 = nkhl_05 + Cdbl(rsAttDaily("khl_05"))
			nTotal_05 = nTotal_05 + Cdbl(rsAttDaily("total_05"))
			
			nkbt_06 = nkbt_06 + Cdbl(rsAttDaily("kbt_06"))
			nkht_06 = nkht_06 + Cdbl(rsAttDaily("kht_06"))
			nkhl_06 = nkhl_06 + Cdbl(rsAttDaily("khl_06"))
			nTotal_06 = nTotal_06 + Cdbl(rsAttDaily("total_06"))
			
			nkbt_07 = nkbt_07 + Cdbl(rsAttDaily("kbt_07"))
			nkht_07 = nkht_07 + Cdbl(rsAttDaily("kht_07"))
			nkhl_07 = nkhl_07 + Cdbl(rsAttDaily("khl_07"))
			nTotal_07 = nTotal_07 + Cdbl(rsAttDaily("total_07"))
			
			nkbt_08 = nkbt_08 + Cdbl(rsAttDaily("kbt_08"))
			nkht_08 = nkht_08 + Cdbl(rsAttDaily("kht_08"))
			nkhl_08 = nkhl_08 + Cdbl(rsAttDaily("khl_08"))
			nTotal_08 = nTotal_08 + Cdbl(rsAttDaily("total_08"))
			
			nkbt_09 = nkbt_09 + Cdbl(rsAttDaily("kbt_09"))
			nkht_09 = nkht_09 + Cdbl(rsAttDaily("kht_09"))
			nkhl_09 = nkhl_09 + Cdbl(rsAttDaily("khl_09"))
			nTotal_09 = nTotal_09 + Cdbl(rsAttDaily("total_09"))
			
			nkbt_10 = nkbt_10 + Cdbl(rsAttDaily("kbt_10"))
			nkht_10 = nkht_10 + Cdbl(rsAttDaily("kht_10"))
			nkhl_10 = nkhl_10 + Cdbl(rsAttDaily("khl_10"))
			nTotal_10 = nTotal_10 + Cdbl(rsAttDaily("total_10"))
			
			nkbt_11 = nkbt_11 + Cdbl(rsAttDaily("kbt_11"))
			nkht_11 = nkht_11 + Cdbl(rsAttDaily("kht_11"))
			nkhl_11 = nkhl_11 + Cdbl(rsAttDaily("khl_11"))
			nTotal_11 = nTotal_11 + Cdbl(rsAttDaily("total_11"))
			
			nkbt_12 = nkbt_12 + Cdbl(rsAttDaily("kbt_12"))
			nkht_12 = nkht_12 + Cdbl(rsAttDaily("kht_12"))
			nkhl_12 = nkhl_12 + Cdbl(rsAttDaily("khl_12"))
			nTotal_12 = nTotal_12 + Cdbl(rsAttDaily("total_12"))
			
			nkbt_SBI = nkbt_SBI + Cdbl(rsAttDaily("kbt_SBI"))
			nkht_SBI = nkht_SBI + Cdbl(rsAttDaily("kht_SBI"))
			nkhl_SBI = nkhl_SBI + Cdbl(rsAttDaily("khl_SBI"))
			nTotal_SBI = nTotal_SBI + cdbl(total_SBI)
			
			nTotal_BUD = nTotal_BUD + cdbl(rsAttDaily("bud_SBI"))
			rsAttDaily.MoveNext
		Wend
	end if
	rsAttDaily.Close
	set rsAttDaily = Nothing
%>
			<tr bgcolor="#F4F4A0">
				<td align="center" colspan="<%= 1 + nSpanTotal %>">TOTAL</td>
				<td align="right"><%=FormatNumber(nkbt_01,2)%></td>
				<td align="right"><%=FormatNumber(nkht_01,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_01,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_01,2)%></td>	
				
				<td align="right"><%=FormatNumber(nkbt_02,2)%></td>
				<td align="right"><%=FormatNumber(nkht_02,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_02,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_02,2)%></td>
				
				<td align="right"><%=FormatNumber(nkbt_03,2)%></td>
				<td align="right"><%=FormatNumber(nkht_03,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_03,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_03,2)%></td>
				
				<td align="right"><%=FormatNumber(nkbt_04,2)%></td>
				<td align="right"><%=FormatNumber(nkht_04,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_04,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_04,2)%></td>
				
				<td align="right"><%=FormatNumber(nkbt_05,2)%></td>
				<td align="right"><%=FormatNumber(nkht_05,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_05,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_05,2)%></td>
				
				<td align="right"><%=FormatNumber(nkbt_06,2)%></td>
				<td align="right"><%=FormatNumber(nkht_06,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_06,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_06,2)%></td>
				
				<td align="right"><%=FormatNumber(nkbt_07,2)%></td>
				<td align="right"><%=FormatNumber(nkht_07,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_07,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_07,2)%></td>
				
				<td align="right"><%=FormatNumber(nkbt_08,2)%></td>
				<td align="right"><%=FormatNumber(nkht_08,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_08,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_08,2)%></td>
				
				<td align="right"><%=FormatNumber(nkbt_09,2)%></td>
				<td align="right"><%=FormatNumber(nkht_09,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_09,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_09,2)%></td>
				
				<td align="right"><%=FormatNumber(nkbt_10,2)%></td>
				<td align="right"><%=FormatNumber(nkht_10,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_10,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_10,2)%></td>
				
				<td align="right"><%=FormatNumber(nkbt_11,2)%></td>
				<td align="right"><%=FormatNumber(nkht_11,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_11,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_11,2)%></td>
				
				<td align="right"><%=FormatNumber(nkbt_12,2)%></td>
				<td align="right"><%=FormatNumber(nkht_12,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_12,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_12,2)%></td>
				
				<td align="right"><%=FormatNumber(nkbt_SBI,2)%></td>
				<td align="right"><%=FormatNumber(nkht_SBI,2)%></td>
				<td align="right"><%=FormatNumber(nkhl_SBI,2)%></td>
				<td align="right"><%=FormatNumber(nTotal_SBI,2)%></td>
				
				<td align="right"><%=FormatNumber(nTotal_BUD,2)%></td>
				
			</tr>
<tbody>
</table>
</form>

</div>
