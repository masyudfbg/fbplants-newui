<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
	Dim v_txtYear, v_txtMonth, v_FirstDate, v_LastDate
	Dim ScriptName
	Dim rsMonitoring, strMonitoring
	Dim val_loop, ColorID
	Dim CurrModule, PrevModule, CurrSubModule, PrevSubModule
	Dim CountallField
	
	Dim val_Title : val_Title = "Laporan Monitoring Employee dan Vehicle Non Activity"
	
	ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

	v_txtYear		=	Request("txtYear")
	v_txtMonth		=	Request("txtMonth")

	If Len(Trim(v_txtYear)) = 0 Then
		v_txtYear = Year(Session("v_Costbook"))
	End If

	If Len(Trim(v_txtMonth)) = 0 Then
		v_txtMonth = Month(Session("v_Costbook"))
	End If
	
	v_FirstDate = Lpad(v_txtMonth, "0", "2") & "/01/" & v_txtYear
	v_LastDate = Lpad(v_txtMonth, "0", "2") & "/" & Lpad(GetLastDay(v_txtMonth, v_txtYear), "0", "2") & "/" & v_txtYear
	
	strMonitoring		=	" Select 'EMPLOYEE' modules" & _
							"        , fccode" & _
							"        , fcname" & _
							"        , detail" & _
							"        , status" & _
							"        , sum(day_1) day_1" & _
							"        , sum(day_2) day_2" & _
							"        , sum(day_3) day_3" & _
							"        , sum(day_4) day_4" & _
							"        , sum(day_5) day_5" & _
							"        , sum(day_6) day_6" & _
							"        , sum(day_7) day_7" & _
							"        , sum(day_8) day_8" & _
							"        , sum(day_9) day_9" & _
							"        , sum(day_10) day_10" & _
							"        , sum(day_11) day_11" & _
							"        , sum(day_12) day_12" & _
							"        , sum(day_13) day_13" & _
							"        , sum(day_14) day_14" & _
							"        , sum(day_15) day_15" & _
							"        , sum(day_16) day_16" & _
							"        , sum(day_17) day_17" & _
							"        , sum(day_18) day_18" & _
							"        , sum(day_19) day_19" & _
							"        , sum(day_20) day_20" & _
							"        , sum(day_21) day_21" & _
							"        , sum(day_22) day_22" & _
							"        , sum(day_23) day_23" & _
							"        , sum(day_24) day_24" & _
							"        , sum(day_25) day_25" & _
							"        , sum(day_26) day_26" & _
							"        , sum(day_27) day_27" & _
							"        , sum(day_28) day_28" & _
							"        , sum(day_29) day_29" & _
							"        , sum(day_30) day_30" & _
							"        , sum(day_31) day_31" & _
							" from (" & _
							"     Select fccode" & _
							"            , fcname" & _
							"            , detail" & _
							"            , status" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 1, datainput, 0)) day_1" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 2, datainput, 0)) day_2" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 3, datainput, 0)) day_3" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 4, datainput, 0)) day_4" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 5, datainput, 0)) day_5" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 6, datainput, 0)) day_6" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 7, datainput, 0)) day_7" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 8, datainput, 0)) day_8" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 9, datainput, 0)) day_9" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 10, datainput, 0)) day_10" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 11, datainput, 0)) day_11" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 12, datainput, 0)) day_12" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 13, datainput, 0)) day_13" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 14, datainput, 0)) day_14" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 15, datainput, 0)) day_15" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 16, datainput, 0)) day_16" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 17, datainput, 0)) day_17" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 18, datainput, 0)) day_18" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 19, datainput, 0)) day_19" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 20, datainput, 0)) day_20" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 21, datainput, 0)) day_21" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 22, datainput, 0)) day_22" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 23, datainput, 0)) day_23" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 24, datainput, 0)) day_24" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 25, datainput, 0)) day_25" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 26, datainput, 0)) day_26" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 27, datainput, 0)) day_27" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 28, datainput, 0)) day_28" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 29, datainput, 0)) day_29" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 30, datainput, 0)) day_30" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 31, datainput, 0)) day_31" & _
							"     from (" & _
							"         select to_date (today, 'mm/dd/yyyy') fddate" & _
							"                , fccode" & _
							"                , fcname" & _
							"                , gangcode detail" & _
							"                , presentgrade status" & _
							"                , 0 datainput" & _
							"         from ( select to_char (to_date ('" & v_FirstDate & "', 'mm/dd/yyyy') + rownum - 1, 'mm/dd/yyyy') today" & _
							"                from all_objects" & _
							"                where rownum <= to_date ('" & v_LastDate & "', 'mm/dd/yyyy') - to_date ('" & v_FirstDate & "', 'mm/dd/yyyy') + 1" & _
							"         ) , employee" & _
							"         where fcba = '" & Session("fcba") & "'" & _
							"               and dateterminate is null" & _
							"         union all" & _
							"         select ag.fddate" & _
							"                , ag.employeecode fccode" & _
							"                , em.fcname" & _
							"                , em.gangcode detail" & _
							"                , em.presentgrade status" & _
							"                , 1 datainput" & _
							"         from attendance_gad ag" & _
							"              , employee em" & _
							"         where ag.fcba = em.fcba" & _
							"               and ag.employeecode = em.fccode" & _
							"               and ag.fcba = '" & Session("fcba") & "'" & _
							"               and em.dateterminate is null" & _
							"               and fddate >= to_date ('" & v_FirstDate & "', 'mm/dd/yyyy')" & _
							"               and fddate <= to_date ('" & v_LastDate & "', 'mm/dd/yyyy')" & _
							"         group by ag.fddate" & _
							"                  , ag.employeecode" & _
							"                  , em.fcname" & _
							"                  , em.gangcode" & _
							"                  , em.presentgrade " & _
							"        UNION ALL " & _
							"         select DISTINCT b.bkmdate fddate" & _
							"                , a.employeecode fccode" & _
							"                , employee_api.get_employeename(a.fcba,a.employeecode) fcname" & _
							"                , b.gangcode detail" & _
							"                , employee_api.get_presentgrade(a.fcba,a.employeecode) status" & _
							"                , 1 datainput" & _
							"         from bkmupahdtl a inner join bkmhdr b on (a.bkmno=b.bkmno and a.fcba=b.fcba) " & _
							"         where b.postingflg='1' and b.fcba = '" & Session("fcba") & "'" & _
							"               and b.bkmdate >= to_date ('" & v_FirstDate & "', 'mm/dd/yyyy')" & _
							"               and b.bkmdate <= to_date ('" & v_LastDate & "', 'mm/dd/yyyy')" & _
							"     )" & _
							"     group by fccode" & _
							"              , fcname" & _
							"              , detail" & _
							"              , status" & _
							" )" & _
							" group by fccode" & _
							"          , fcname" & _
							"          , detail" & _
							"          , status" & _
							" having sum(day_1)" & _
							"        + sum(day_2)" & _
							"        + sum(day_3)" & _
							"        + sum(day_4)" & _
							"        + sum(day_5)" & _
							"        + sum(day_6)" & _
							"        + sum(day_7)" & _
							"        + sum(day_8)" & _
							"        + sum(day_9)" & _
							"        + sum(day_10)" & _
							"        + sum(day_11)" & _
							"        + sum(day_12)" & _
							"        + sum(day_13)" & _
							"        + sum(day_14)" & _
							"        + sum(day_15)" & _
							"        + sum(day_16)" & _
							"        + sum(day_17)" & _
							"        + sum(day_18)" & _
							"        + sum(day_19)" & _
							"        + sum(day_20)" & _
							"        + sum(day_21)" & _
							"        + sum(day_22)" & _
							"        + sum(day_23)" & _
							"        + sum(day_24)" & _
							"        + sum(day_25)" & _
							"        + sum(day_26)" & _
							"        + sum(day_27)" & _
							"        + sum(day_28)" & _
							"        + sum(day_29)" & _
							"        + sum(day_30)" & _
							"        + sum(day_31)" & _
							"        = 0" & _
							" Union All" & _
							" Select 'VEHICLE' modules" & _
							"        , fccode" & _
							"        , fcname" & _
							"        , detail" & _
							"        , status" & _
							"        , sum(day_1) day_1" & _
							"        , sum(day_2) day_2" & _
							"        , sum(day_3) day_3" & _
							"        , sum(day_4) day_4" & _
							"        , sum(day_5) day_5" & _
							"        , sum(day_6) day_6" & _
							"        , sum(day_7) day_7" & _
							"        , sum(day_8) day_8" & _
							"        , sum(day_9) day_9" & _
							"        , sum(day_10) day_10" & _
							"        , sum(day_11) day_11" & _
							"        , sum(day_12) day_12" & _
							"        , sum(day_13) day_13" & _
							"        , sum(day_14) day_14" & _
							"        , sum(day_15) day_15" & _
							"        , sum(day_16) day_16" & _
							"        , sum(day_17) day_17" & _
							"        , sum(day_18) day_18" & _
							"        , sum(day_19) day_19" & _
							"        , sum(day_20) day_20" & _
							"        , sum(day_21) day_21" & _
							"        , sum(day_22) day_22" & _
							"        , sum(day_23) day_23" & _
							"        , sum(day_24) day_24" & _
							"        , sum(day_25) day_25" & _
							"        , sum(day_26) day_26" & _
							"        , sum(day_27) day_27" & _
							"        , sum(day_28) day_28" & _
							"        , sum(day_29) day_29" & _
							"        , sum(day_30) day_30" & _
							"        , sum(day_31) day_31" & _
							" from (" & _
							"     Select fccode" & _
							"            , fcname" & _
							"            , detail" & _
							"            , status" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 1, datainput, 0)) day_1" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 2, datainput, 0)) day_2" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 3, datainput, 0)) day_3" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 4, datainput, 0)) day_4" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 5, datainput, 0)) day_5" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 6, datainput, 0)) day_6" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 7, datainput, 0)) day_7" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 8, datainput, 0)) day_8" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 9, datainput, 0)) day_9" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 10, datainput, 0)) day_10" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 11, datainput, 0)) day_11" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 12, datainput, 0)) day_12" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 13, datainput, 0)) day_13" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 14, datainput, 0)) day_14" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 15, datainput, 0)) day_15" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 16, datainput, 0)) day_16" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 17, datainput, 0)) day_17" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 18, datainput, 0)) day_18" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 19, datainput, 0)) day_19" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 20, datainput, 0)) day_20" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 21, datainput, 0)) day_21" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 22, datainput, 0)) day_22" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 23, datainput, 0)) day_23" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 24, datainput, 0)) day_24" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 25, datainput, 0)) day_25" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 26, datainput, 0)) day_26" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 27, datainput, 0)) day_27" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 28, datainput, 0)) day_28" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 29, datainput, 0)) day_29" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 30, datainput, 0)) day_30" & _
							"            , sum(decode(to_number(to_char(fddate,'dd')), 31, datainput, 0)) day_31" & _
							"     from (" & _
							"         select to_date (today, 'mm/dd/yyyy') fddate" & _
							"                , fccode" & _
							"                , fcname" & _
							"                , vg.description detail" & _
							"                , to_char(vh.yearofpurchase, 'yyyy') status" & _
							"                , 0 datainput" & _
							"         from ( select to_char (to_date ('" & v_FirstDate & "', 'mm/dd/yyyy') + rownum - 1, 'mm/dd/yyyy') today" & _
							"                from all_objects" & _
							"                where rownum <= to_date ('" & v_LastDate & "', 'mm/dd/yyyy') - to_date ('" & v_FirstDate & "', 'mm/dd/yyyy') + 1" & _
							"         ) , vehicle vh" & _
							"           , vehiclegroup vg" & _
							"         where vh.fcba = vg.fcba" & _
							"               and vh.vehiclegroupcode = vg.vehiclegroupcode" & _
							"               and vh.fcba = '" & Session("fcba") & "'" & _
							"               and dateterminate is null" & _
							"               and activation = 'Y'" & _
							"         union all" & _
							"         select va.vehicledate" & _
							"                , va.vehiclecode fccode" & _
							"                , vh.fcname" & _
							"                , vg.description detail" & _
							"                , to_char(vh.yearofpurchase, 'yyyy') status" & _
							"                , 1 datainput" & _
							"         from vehicleactivity va" & _
							"              , vehicle vh" & _
							"              , vehiclegroup vg" & _
							"         where va.fcba = vh.fcba" & _
							"               and va.vehiclecode = vh.fccode" & _
							"               and vh.fcba = vg.fcba" & _
							"               and vh.vehiclegroupcode = vg.vehiclegroupcode" & _
							"               and va.fcba = '" & Session("fcba") & "'" & _
							"               and vh.dateterminate is null" & _
							"               and vh.activation = 'Y'" & _
							"               and vehicledate >= to_date ('" & v_FirstDate & "', 'mm/dd/yyyy')" & _
							"               and vehicledate <= to_date ('" & v_LastDate & "', 'mm/dd/yyyy')" & _
							"         group by va.vehicledate" & _
							"                  , va.vehiclecode" & _
							"                  , vh.fcname" & _
							"                  , vg.description" & _
							"                  , to_char(vh.yearofpurchase, 'yyyy')" & _
							"         union all" & _
							"         select DISTINCT b.bkmdate vehicledate" & _
							"                , b.vehiclecode fccode" & _
							"                , vehicle_api.get_vehiclecode_name(b.vehiclecode,b.fcba) fcname" & _
							"                , vehiclegroup_api.get_vehiclegroup_name(vehicle_api.get_vehiclegroupcode(b.vehiclecode,b.fcba),b.fcba) detail" & _
							"                , vehicle_api.get_vehicle_year(b.vehiclecode,b.fcba) status" & _
							"                , 1 datainput" & _
							"         from bkmvehicledtl a inner join bkmhdr b on( a.bkmno=b.bkmno and a.fcba=b.fcba)" & _
							"         where b.postingflg='1' and b.fcba = '" & Session("fcba") & "'" & _
							"               and bkmdate >= to_date ('" & v_FirstDate & "', 'mm/dd/yyyy')" & _
							"               and bkmdate <= to_date ('" & v_LastDate & "', 'mm/dd/yyyy')" & _
							"     )" & _
							"     group by fccode" & _
							"              , fcname" & _
							"              , detail" & _
							"              , status" & _
							" )" & _
							" group by fccode" & _
							"          , fcname" & _
							"          , detail" & _
							"          , status" & _
							" having sum(day_1)" & _
							"        + sum(day_2)" & _
							"        + sum(day_3)" & _
							"        + sum(day_4)" & _
							"        + sum(day_5)" & _
							"        + sum(day_6)" & _
							"        + sum(day_7)" & _
							"        + sum(day_8)" & _
							"        + sum(day_9)" & _
							"        + sum(day_10)" & _
							"        + sum(day_11)" & _
							"        + sum(day_12)" & _
							"        + sum(day_13)" & _
							"        + sum(day_14)" & _
							"        + sum(day_15)" & _
							"        + sum(day_16)" & _
							"        + sum(day_17)" & _
							"        + sum(day_18)" & _
							"        + sum(day_19)" & _
							"        + sum(day_20)" & _
							"        + sum(day_21)" & _
							"        + sum(day_22)" & _
							"        + sum(day_23)" & _
							"        + sum(day_24)" & _
							"        + sum(day_25)" & _
							"        + sum(day_26)" & _
							"        + sum(day_27)" & _
							"        + sum(day_28)" & _
							"        + sum(day_29)" & _
							"        + sum(day_30)" & _
							"        + sum(day_31)" & _
							"        = 0" & _
							" ORDER BY modules,detail,status,fcname "
							
							


	'Response.Write strMonitoring
	
	Set rsMonitoring = Server.CreateObject("ADODB.Recordset")
	rsMonitoring.Open strMonitoring, connstr, adOpenStatic 
	
%>


<html>

<head>
<title>ePlaS Reporting - <%=PCase(val_Title)%></title>
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
<thead>
	<tr>
		<th colspan="1" style="background-color:#FFFFFF; border:none;">  
			<div id="header-wrapper">
				<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=Session("COMPANY")%></h4>
					<h5><%=Session("BA")%></h5>
				</div>
		</div>
		</th>
		<th colspan="35" style="background-color:#FFFFFF; border:none;">
		  <div id="subheader" class="centered" tyle="width:50%; margin:auto;">
					<h2><%=Ucase(val_Title)%></h2>
					<br/>
					<h4>PERIODE&nbsp;:&nbsp;<%=textMonth(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
				</div>
		</th>
	</tr>
	<tr>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">KODE</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">NAMA</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">DETAIL</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">STATUS</th>
		<th colspan="31" style="background-color:#00A300;color:#fff;">TANGGAL</th>
	</tr>
	<tr>
			<%For val_loop = 1 To 31%>
				<th style="background-color:#00A300;color:#fff;"><%=val_loop%></th>
			<%Next%>
	</tr>
</thead>
<tbody>    
<%If rsMonitoring.BOF and rsMonitoring.EOF then%>
		<tr>
			<td colspan="35" align="left">NO DATA AVAILABLE</td>
		</tr>
<%Else%>
	<%CountallField = rsMonitoring.fields.count-1%>
	<%while not rsMonitoring.eof%>
		<%CurrModule = rsMonitoring("modules")%>
		<%If CurrModule <> PrevModule Then%>
			<tr>
				<td colspan="35" align="left" style="padding-left:10px;" bgcolor="#F4F4A0"><%=CurrModule%></td>
			</tr>
			<tr>
				<td align="left"><%=rsMonitoring("fccode")%></td>
				<td align="left"><%=rsMonitoring("fcname")%></td>
				<td align="left"><%=rsMonitoring("detail")%></td>
				<td align="center"><%=rsMonitoring("status")%></td>
				<%For val_loop = 5 to CountallField%>
					<%If Cdbl(rsMonitoring(val_loop)) = 1 Then
						ColorID = "<img alt="""" src=""/images/check.gif""/>"
					Else
						ColorID = "&nbsp;"
					End If%>
					<td align="right" width="1%"><%=ColorID%></td>
				<%Next%>
			</tr>
		<%Else%>
			<tr>
				<td align="left"><%=rsMonitoring("fccode")%></td>
				<td align="left"><%=rsMonitoring("fcname")%></td>
				<td align="left"><%=rsMonitoring("detail")%></td>
				<td align="center"><%=rsMonitoring("status")%></td>
				<%For val_loop = 5 to CountallField%>
					<%If Cdbl(rsMonitoring(val_loop)) = 1 Then
						ColorID = "<img alt="""" src=""/images/check.gif""/>"
					Else
						ColorID = "&nbsp;"
					End If%>
					<td align="right" width="1%"><%=ColorID%></td>
				<%Next%>
			</tr>
		<%End If%>
<%
			PrevModule = CurrModule
			PrevSubModule = CurrSubModule
			rsMonitoring.MoveNext
		Wend
End If
strMonitoring = ""
rsMonitoring.Close
set rsMonitoring = Nothing
%>

	</tbody>
</table>

<!--#include file="../Bottom_Report.asp"-->