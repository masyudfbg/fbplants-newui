<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtYear, v_txtMonth, v_txtMonth1, no
Dim ScriptName
Dim rsgaji, strgaji, strMain, strgaji2
Dim cClass, nday, xNum, nSpanTotal, FirstDate

Dim var_colspan : var_colspan = 1

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType

nSpanTotal = 7

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
					" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba, BUSINESSUNIT_API.get_fcname(fcba) fcbaname"
					
wherecons	=	" where company is not null "


If Len(Trim(v_txtGroupPlantation)) > 0 Then
	if v_txtGroupPlantation = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
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
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtPlantation = "--ALL--" then
		selectCons	=	selectCons	& " , Plantation"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , Plantation"
		whereCons	=	whereCons	& " And Plantation = '" & v_txtPlantation & "'"
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_colspan	=	var_colspan	+ 1
	End If
End If


If Len(Trim(v_txtGroupRegion)) > 0 Then
	if v_txtGroupRegion = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
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
		nSpanTotal	=	nSpanTotal	- 1
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
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtCompany = "--ALL--" then
		selectCons	=	selectCons	& " , company"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , company"
		whereCons	=	whereCons	& " And company = '" & v_txtCompany & "'"
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
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
		nSpanTotal	=	nSpanTotal	- 2
	Elseif v_txtfcba = "--ALL--" then
		selectCons	=	selectCons	& " , nourut, fcba, fcbaname"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , nourut, fcba "
		var_colspan	=	var_colspan	+ 2
	Else
		selectCons	=	selectCons	& " , nourut, fcba, fcbaname"
		whereCons	=	whereCons	& " And fcba = '" & v_txtfcba & "'"
		orderCons	=	orderCons	& " , nourut"
		var_colspan	=	var_colspan	+ 2
		
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

v_txtYear		= Request("txtYear")
v_txtMonth		= Request("txtMonth")
v_txtMonth1		= Request("txtMonth1")

If Len(Trim(v_txtYear)) = 0 Then
	v_txtYear	= Year(Session("v_Costbook"))
End If
If Len(Trim(v_txtMonth)) = 0 Then
	v_txtMonth	= Month(Session("v_Costbook"))
End If
If Len(Trim(v_txtMonth1)) = 0 Then
	v_txtMonth1	= Month(Session("v_Costbook"))
End If

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain		=	" select" & strFieldCons & _
				"        , sum(base_salary) base_salary" & _
				"        , sum(rapelgaji) rapelgaji" & _
				"        , sum(overtime) overtime" & _
				"        , sum(premi) premi" & _
				"        , sum(tunjangan) tunjangan" & _
				"        , sum(borongan) borongan" & _
				"        , sum(p_jamsostek) p_jamsostek" & _
				"        , SUM(P_BPJSKES) P_BPJSKES" & _
				"        , sum(p_mk_p1) p_mk_p1" & _
				"        , sum(p_panen) p_panen" & _
				"        , sum(p_pph21) p_pph21" & _
				"        , SUM(P_KOPERASI) P_KOPERASI" & _
				"        , sum(p_lain) p_lain" & _
				"        , sum(total) total" & _
				"        , sum(decode(paymenttype, 'TRANSFER', total, 0)) transfer" & _
				"        , sum(decode(paymenttype, 'NONTRANSFER', total, 0)) nontransfer" & _
				" from (" & _
				"     select fcba" & _
				"            , fccode" & _
				"            , nvl(employee_api.get_paymenttype(fcba, fccode), 'NONTRANSFER') paymenttype" & _
				"            , sum(base_salary) base_salary" & _
				"            , sum(rapelgaji) rapelgaji" & _
				"            , sum(overtime) overtime" & _
				"            , sum(premi) premi" & _
				"            , sum(tunjangan) tunjangan" & _
				"            , sum(borongan) borongan" & _
				"            , sum(p_jamsostek) p_jamsostek" & _
				"            , SUM(P_BPJSKES) P_BPJSKES" & _
				"            , sum(p_mk_p1) p_mk_p1" & _
				"            , sum(p_panen) p_panen" & _
				"            , sum(p_pph21) p_pph21" & _
				"            , SUM(P_KOPERASI) P_KOPERASI" & _
				"            , sum(p_lain) p_lain" & _
				"            , sum(base_salary)" & _
				"              + sum(rapelgaji)" & _
				"              + sum(overtime)" & _
				"              + sum(premi)" & _
				"              + sum(tunjangan)" & _
				"              + sum(borongan)" & _
				"              + sum(p_jamsostek)" & _
				"              + SUM(P_BPJSKES)" & _
				"              + sum(p_mk_p1)" & _
				"              + sum(p_panen)" & _
				"              + sum(p_pph21)" & _
				"              + SUM(P_KOPERASI)" & _
				"              + sum(p_lain) total" & _
				"     from view_payrolldetail" & _
				"     where to_number(to_char(substr(yearmonth, 1, 4))) = '" & v_txtYear & "'" & _
				"           and to_number(to_char(substr(yearmonth, 5, 2))) between '" & v_txtMonth & "' and '" & v_txtMonth1 & "'" & _
				"     group by fcba" & _
				"              , fccode" & _
				"              , nvl(employee_api.get_paymenttype(fcba, fccode), 'NONTRANSFER')" & _
				"	   HAVING SUM (base_salary) + SUM (overtime)+ SUM (borongan) + SUM (premi) > 0 " & _ 
				" UNION " & _
				"					SELECT " & _ 
				"                              t.fcba,"  & _
				"                              t.employeecode fccode,"  & _
				"            				   nvl(employee_api.get_paymenttype(t.fcba, t.employeecode), 'NONTRANSFER') paymenttype," & _
				"                              SUM (gaji) base_salary,"  & _
				"                              0 rapelgaji," & _
				"                              SUM (t.lembur) overtime,"  & _
				"                              SUM (t.premi) premi,"  & _
				"                              SUM (t.tunjangan) tunjangan,"  & _
				"                              SUM (t.borongan) borongan,"  & _
				"                              SUM (t.pot_bpjs_tk) p_jamsostek,"  & _
				"                              SUM (t.pot_bpjs_ks) p_bpjskes,"  & _
				"                              SUM (0) p_mk_p1,"  & _
				"                              SUM (t.pot_denda) p_panen," & _
				"                              SUM (0) p_pph21,"  & _
				"                              SUM (0) P_KOPERASI," & _
				"                              SUM (t.pot_pinjaman) p_lain,"  & _
				"                              SUM (t.total) total"  & _
				"                    FROM bkmpayroll t"  & _
				"                    WHERE t.period_year = '"& v_txtYear &"'"  & _
				"                              AND TO_NUMBER(t.period_month) between '"& v_txtMonth &"' and '"& v_txtMonth1 &"' "  & _
				"                     GROUP BY "  & _
				"                              t.fcba,"  & _
				"                              t.employeecode, "  & _
				"                              nvl(employee_api.get_paymenttype(t.fcba, t.employeecode), 'NONTRANSFER')" & _
				"                     HAVING SUM (t.subgaji) > 0 " & _ 
				
				" )" & _
				" group by fcba"

				
Dim Selection_2

Selection_2 = " "

strgaji2	=   "        , sum(base_salary) base_salary" & _
				"        , sum(rapelgaji) rapelgaji" & _
				"        , sum(overtime) overtime" & _
				"        , sum(premi) premi" & _
				"        , sum(tunjangan) tunjangan" & _
				"        , sum(borongan) borongan" & _
				"        , sum(p_jamsostek) p_jamsostek" & _
				"        , sum(P_BPJSKES) P_BPJSKES" & _
				"        , sum(p_mk_p1) p_mk_p1" & _
				"        , sum(p_panen) p_panen" & _
				"        , sum(p_pph21) p_pph21" & _
				"        , sum(P_KOPERASI) P_KOPERASI" & _
				"        , sum(p_lain) p_lain" & _
				"        , sum(total) total" & _
                "        , sum(transfer) transfer" & _
				"        , sum(nontransfer) nontransfer"

strgaji	=	" select GroupPlantation " & selectCons & Selection_2 & strgaji2 & "  From (" & strMain & " )" & whereCons & _
					" group by GroupPlantation " & selectCons & Selection_2 & " Order by GroupPlantation " & orderCons & " "

'response.write strgaji
%>
<html>

<head>
<title>ePlaS Reporting - Laporan Rekap Gaji</title>
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
				<!--#include virtual="/REPORT/FormSelectNewSum.html"-->
				<table id="form-table" class="centered-margin">
					<tr>
						<td>
							<font face="verdana" size="1">TAHUN</font>
						</td>
						<% 
							Dim mon, Yr, MinYr, MaxYr, i, MM
							mon = Month(Session("v_Costbook"))
							If len(v_txtMonth)=0 then v_txtMonth=mon
							If len(v_txtMonth1)=0 then v_txtMonth1=mon
						
							Yr = Year(Session("v_Costbook"))
							MinYr = 2015
							MaxYr = Yr  + 3
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
							<font face="verdana" size="1">DARI BULAN</font>
						</td>
						<td width="2">
							<select name="txtMonth">
							<%For i = 1 to 12%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
							<%Next%>
							</select>
						</td>
						<td>
							<font face="verdana" size="1">SAMPAI BULAN</font>
						</td>
						<td width="2">
							<select name="txtMonth1">
							<%For i = 1 to 12%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth1) then%> Selected <%End If%>><%=Bulan(i)%></option>
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
      <table border="0" width="100%" id="product-table" class="xthf-site">
	<thead style="border-left:1px #000000 solid;">
        <tr>
            <th colspan="<%=var_colspan+15%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
            <div id="header">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4>
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>
			<% End if%>
            </div>
            <div id="subheader" class="centered">
                <h2>REKAP GAJI KARYAWAN</h3>
                <br/>
                <h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%> s/d <%=Bulan(v_txtMonth1)%>&nbsp;<%=v_txtYear%></h4>
            </div>
		</div>

</th>
</tr>
	<tr>
		
		<%If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleGroupPlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtPlantation)) > 0 And v_txtPlantation <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtCompany)) > 0 And v_txtCompany <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
		<%End If%>

		<%If Len(Trim(v_txtfcba)) > 0 And v_txtfcba <>  "--SUMMARY--" Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">Nama <%=titleSite%></th>
		<%End If%>
		
		<%If var_colspan = 1 Then%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">&nbsp;</th>
		<%End If%>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">Gaji</th>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">Rapel Gaji</th>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">Lembur</th>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">Premi</th>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">Tunjangan</th>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">Borongan</th>
			<th colspan="7" style="background-color:#00A300;color:#fff;">Potongan</th>
			<th rowspan="2" style="background-color:#00A300;color:#fff;">Total</th>
			<th colspan="2" style="background-color:#00A300;color:#fff;">Dana</th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;">BPJS Ketenagakerjaan</th>
		<th style="background-color:#00A300;color:#fff;">BPJS</th>
		<th style="background-color:#00A300;color:#fff;">MK/P1</th>
		<th style="background-color:#00A300;color:#fff;">Panen</th>
		<th style="background-color:#00A300;color:#fff;">PPh 21</th>
		<th style="background-color:#00A300;color:#fff;">Koperasi</th>
		<th style="background-color:#00A300;color:#fff;">Lain-lain</th>
		<th style="background-color:#00A300;color:#fff;">Transfer</th>
		<th style="background-color:#00A300;color:#fff;">Non Transfer</th>
	</tr>
    </thead>
<tbody>
<%
Set rsgaji = Server.CreateObject("ADODB.Recordset")
rsgaji.Open strgaji, connstr, adOpenStatic

Dim subtotalbase_salary : subtotalbase_salary = 0
Dim subtotalrapelgaji : subtotalrapelgaji = 0
Dim subtotalovertime : subtotalovertime = 0
Dim subtotalpremi : subtotalpremi = 0
Dim subtotaltunjangan : subtotaltunjangan = 0
Dim subtotalborongan : subtotalborongan = 0
Dim subtotalp_jamsostek : subtotalp_jamsostek = 0
Dim subtotalp_bpjs : subtotalp_bpjs = 0
Dim subtotalp_mk_p1 : subtotalp_mk_p1 = 0
Dim subtotalp_panen : subtotalp_panen = 0
Dim subtotalp_pph21 : subtotalp_pph21 = 0
Dim subtotalp_koperasi : subtotalp_koperasi = 0
Dim subtotalp_lain : subtotalp_lain = 0
Dim subtotaltotal : subtotaltotal = 0
Dim subtotaltransfer : subtotaltransfer = 0
Dim subtotalnontransfer : subtotalnontransfer = 0

Dim grandtotalbase_salary : grandtotalbase_salary = 0
Dim grandtotalrapelgaji : grandtotalrapelgaji = 0
Dim grandtotalovertime : grandtotalovertime = 0
Dim grandtotalpremi : grandtotalpremi = 0
Dim grandtotaltunjangan : grandtotaltunjangan = 0
Dim grandtotalborongan : grandtotalborongan = 0
Dim grandtotalp_jamsostek : grandtotalp_jamsostek = 0
Dim grandtotalp_bpjs : grandtotalp_bpjs = 0
Dim grandtotalp_mk_p1 : grandtotalp_mk_p1 = 0
Dim grandtotalp_panen : grandtotalp_panen = 0
Dim grandtotalp_pph21 : grandtotalp_pph21 = 0
Dim grandtotalp_koperasi : grandtotalp_koperasi = 0
Dim grandtotalp_lain : grandtotalp_lain = 0
Dim grandtotaltotal : grandtotaltotal = 0
Dim grandtotaltransfer : grandtotaltransfer = 0
Dim grandtotalnontransfer : grandtotalnontransfer = 0

Dim tempGroupPlantation : tempGroupPlantation = ""
Dim tempPlantation : tempPlantation = ""
Dim tempGroupRegion : tempGroupRegion = ""
Dim tempRegion : tempRegion = ""
Dim tempCompany : tempCompany = ""
Dim tempfcba : tempfcba = ""

Dim SprevGroupPlantation : SprevGroupPlantation = ""
Dim SprevPlantation : SprevPlantation = ""
Dim SprevGroupRegion : SprevGroupRegion = ""
Dim SprevRegion : SprevRegion = ""
Dim SprevCompany : SprevCompany = ""
Dim sprevFcba : sprevFcba = ""

Dim CurrRegion : CurrRegion = ""
Dim PrevRegion : PrevRegion = ""

Dim Currfcba : Currfcba = ""
Dim Prevfcba : Prevfcba = ""

if rsgaji.bof and rsgaji.eof then
%>
	<tr>
		<td colspan="<%=var_colspan+14%>">NO DATA AVAILABLE</td>
	</tr>
<%
else
		while not rsgaji.eof
			if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then
				CurrRegion = rsgaji("region")
				If CurrRegion <> PrevRegion Then
					If Len(Trim(PrevRegion)) = 0 Then
						subtotalbase_salary = 0
						subtotalrapelgaji = 0
						subtotalovertime = 0
						subtotalpremi = 0
						subtotaltunjangan = 0
						subtotalborongan = 0
						subtotalp_jamsostek = 0
						subtotalp_bpjs = 0
						subtotalp_mk_p1 = 0
						subtotalp_panen = 0
						subtotalp_pph21 = 0
						subtotalp_koperasi = 0
						subtotalp_lain = 0
						subtotaltotal = 0
						subtotaltransfer = 0
						subtotalnontransfer = 0
					Else
%>
						<tr style="background-color:#F4F4A0;">
							<td align="left" colspan="<%=var_colspan-1%>" style="padding-left:10px;"><strong>Sub Total <%=PCase(PrevRegion)%></strong></td>
							<td align="Right"><strong><%=FormatNumber(subtotalbase_salary, 0)%></strong></td>
							<td align="Right"><strong><%=FormatNumber(subtotalrapelgaji, 0)%></strong></td>
							<td align="Right"><strong><%=FormatNumber(subtotalovertime, 0)%></strong></td>
							<td align="Right"><strong><%=FormatNumber(subtotalpremi, 0)%></strong></td>
							<td align="Right"><strong><%=FormatNumber(subtotaltunjangan, 0)%></strong></td>
							<td align="Right"><strong><%=FormatNumber(subtotalborongan, 0)%></strong></td>
							<td align="Right"><strong><%=FormatNumber(subtotalp_jamsostek, 0)%></strong></td>
							<td align="Right"><strong><%=FormatNumber(subtotalp_bpjs, 0)%></strong></td>
							<td align="Right"><strong><%=FormatNumber(subtotalp_mk_p1, 0)%></strong></td>
							<td align="Right"><strong><%=FormatNumber(subtotalp_panen, 0)%></strong></td>
							<td align="Right"><strong><%=FormatNumber(subtotalp_pph21, 0)%></strong></td>
							<td align="Right"><strong><%=FormatNumber(subtotalp_koperasi, 0)%></strong></td>
							<td align="Right"><strong><%=FormatNumber(subtotalp_lain, 0)%></strong></td>
							<td align="Right"><strong><%=FormatNumber(subtotaltotal, 0)%></strong></td>
							<td align="Right"><strong><%=FormatNumber(subtotaltransfer, 0)%></strong></td>
							<td align="Right"><strong><%=FormatNumber(subtotalnontransfer, 0)%></strong></td>
						</tr>
<%
						subtotalbase_salary = 0
						subtotalrapelgaji = 0
						subtotalovertime = 0
						subtotalpremi = 0
						subtotaltunjangan = 0
						subtotalborongan = 0
						subtotalp_jamsostek = 0
						subtotalp_bpjs = 0
						subtotalp_mk_p1 = 0
						subtotalp_panen = 0
						subtotalp_pph21 = 0
						subtotalp_koperasi = 0
						subtotalp_lain = 0
						subtotaltotal = 0
						subtotaltransfer = 0
						subtotalnontransfer = 0
%>
					<%End If%>
					<tr>
							
							<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
								<% tempGroupPlantation = rsgaji("GroupPlantation")
								   tempPlantation = "--ALL--"
								   tempGroupRegion = "--SUMMARY--"
								   tempRegion = "--SUMMARY--"
								   tempCompany ="--SUMMARY--"
								   tempfcba = "--SUMMARY--"
								%>
								<% if tempGroupPlantation <> SprevGroupPlantation then %>
									<td align="center">
										<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtMonth1=<%=v_txtMonth1%>">
										<%=rsgaji("GroupPlantation")%>
										</a>
									</td>
									<% SprevGroupPlantation = tempGroupPlantation %>	
								<% else	%>
									<td align="center"></td>					
								<% end if %>
							<% end if %>
							
							
							<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
								<%if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
									  tempGroupPlantation = rsgaji("GroupPlantation")
								   else
									  tempGroupPlantation = ""
								   end if
								
								   tempPlantation = rsgaji("Plantation")
								   tempGroupRegion = "--ALL--"
								   tempRegion = "--SUMMARY--"
								   tempCompany ="--SUMMARY--"
								   tempfcba = "--SUMMARY--"
								%>
								<% if tempPlantation <> SprevPlantation then %>
									<td align="center">
										<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtMonth1=<%=v_txtMonth1%>">
										<%=rsgaji("Plantation")%>
										</a>
									</td>
									<% SprevPlantation = tempPlantation %>	
								<% else	%>
									<td align="center"></td>					
								<% end if %>	
							<% end if %>
							
							<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
								<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
									  tempGroupPlantation = rsgaji("GroupPlantation")
								   else
									  tempGroupPlantation = ""
								   end if					
								   
								   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
									  tempPlantation = rsgaji("Plantation")
								   else
									  tempPlantation = ""
								   end if
								   
								   tempGroupRegion = rsgaji("GroupRegion")
								   tempRegion = "--ALL--"
								   tempCompany ="--SUMMARY--"
								   tempfcba = "--SUMMARY--"
								%>
								<% if tempGroupRegion <> SprevGroupRegion then %>
									<td align="center">
										<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtMonth1=<%=v_txtMonth1%>">
										<%=rsgaji("GroupRegion")%>
										</a>
									</td>
									<% SprevGroupRegion= tempGroupRegion %>	
								<% else	%>
									<td align="center"></td>					
								<% end if %>	
							<% end if %>
							
							<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
								<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
									  tempGroupPlantation = rsgaji("GroupPlantation")
								   else
									  tempGroupPlantation = ""
								   end if					
								   
								   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
									  tempPlantation = rsgaji("Plantation")
								   else
									  tempPlantation = ""
								   end if
								   
								   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
									   tempGroupRegion = rsgaji("GroupRegion")
								   else
									   tempGroupRegion = ""
								   end if
								   
								   tempRegion = rsgaji("Region") 
								   tempCompany ="--ALL--"
								   tempfcba = "--SUMMARY--"
								%>
								
								<% if tempRegion <> SprevRegion then %>
									<td align="center">
										<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtMonth1=<%=v_txtMonth1%>">
										<%=rsgaji("region")%>
										</a>
									</td>
									<% SprevRegion = tempRegion %>	
								<% else	%>
									<td align="center"></td>					
								<% end if %>	
							<% end if %>
							
							<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
								<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
									  tempGroupPlantation = rsgaji("GroupPlantation")
								   else
									  tempGroupPlantation = ""
								   end if					
								   
								   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
									  tempPlantation = rsgaji("Plantation")
								   else
									  tempPlantation = ""
								   end if
								   
								   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
									   tempGroupRegion = rsgaji("GroupRegion")
								   else
									   tempGroupRegion = ""
								   end if
								   
								   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
									   tempRegion = rsgaji("Region") 
								   else
									   tempRegion = ""
								   end if 

								   tempCompany =rsgaji("Company")
								   tempfcba = "--ALL--"	   
								 %>
								
								<% if tempCompany <> SprevCompany then %>
									<td align="center">
										<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtMonth1=<%=v_txtMonth1%>">
										<%=rsgaji("company")%>
										</a>
									</td>
									<% SprevCompany = tempCompany %>	
								<% else	%>
									<td align="center"></td>					
								<% end if %>
							<% end if %>
							
							<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
								<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
									  tempGroupPlantation = rsgaji("GroupPlantation")
								   else
									  tempGroupPlantation = ""
								   end if					
								   
								   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
									  tempPlantation = rsgaji("Plantation")
								   else
									  tempPlantation = ""
								   end if
								   
								   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
									   tempGroupRegion = rsgaji("GroupRegion")
								   else
									   tempGroupRegion = ""
								   end if
								   
								   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
									   tempRegion = rsgaji("Region") 
								   else
									   tempRegion = ""
								   end if 
								   
								   if  v_txtcompany <>  "--SUMMARY--" and v_txtcompany <> "" then
									   tempCompany =rsgaji("Company")
								   else
									   tempCompany =""
								   end if

								   tempfcba = rsgaji("fcba")	   
								 %>
								
								<% if tempfcba <> sprevFcba then %>
									<td align="center">
										<a href="p_rekappayroll_2.asp?txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtMonth1=<%=v_txtMonth1%>">
										<%=rsgaji("fcba")%>
										</a>
									</td>
									<td align="left"><%=rsgaji("fcbaname")%></td>
									<% sprevFcba = tempfcba %>	
								<% else	%>
									<td align="center"></td>
									<td align="center"></td>
								<% end if %>
							<% end if %>

							<%If var_colspan = 1 Then%>
								<td align="right"></td>
							<%End If%>
								<td align="right"><%=FormatNumber(rsgaji("base_salary"), 0)%></td>
								<td align="right"><%=formatnumber( cdbl(rsgaji("rapelgaji")), 0)%></td>
								<td align="right"><%=formatnumber( cdbl(rsgaji("overtime")), 0)%></td>
								<td align="right"><%=formatnumber( cdbl(rsgaji("premi")), 0)%></td>
								<td align="right"><%=formatnumber( cdbl(rsgaji("tunjangan")), 0)%></td>
								<td align="right"><%=formatnumber( cdbl(rsgaji("borongan")), 0)%></td>
								<td align="right"><%=formatnumber( cdbl(rsgaji("p_jamsostek")), 0)%></td>
								<td align="right"><%=formatnumber( cdbl(rsgaji("P_BPJSKES")), 0)%></td>
								<td align="right"><%=formatnumber( cdbl(rsgaji("p_mk_p1")), 0)%></td>
								<td align="right"><%=formatnumber( cdbl(rsgaji("p_panen")), 0)%></td>
								<td align="right"><%=formatnumber( cdbl(rsgaji("p_pph21")), 0)%></td>
								<td align="right"><%=formatnumber( cdbl(rsgaji("P_KOPERASI")), 0)%></td>
								<td align="right"><%=formatnumber( cdbl(rsgaji("p_lain")), 0)%></td>
								<td align="right"><%=formatnumber( cdbl(rsgaji("total")), 0)%></td>
								<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
									<% currfcba = rsgaji("fcba") %>
									<% if currfcba <> prevfcba then %>
										<td align="right">
											<a href="P_RekapTransferGaji.asp?txtfcba=<%=currfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtMonth1=<%=v_txtMonth1%>">
											<%=formatnumber( cdbl(rsgaji("transfer")), 0)%>
											</a>
										</td>
										<% prevfcba = currfcba %>
									<% end if %>
								<%else%>
									<td align="right"><%=formatnumber( cdbl(rsgaji("transfer")), 0)%></td>
								<% end if %>
								<td align="right"><%=formatnumber( cdbl(rsgaji("nontransfer")), 0)%></td>
					</tr>
				<%Else%>
					<tr>
					
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
						<% tempGroupPlantation = rsgaji("GroupPlantation")
						   tempPlantation = "--ALL--"
						   tempGroupRegion = "--SUMMARY--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						%>
						<% if tempGroupPlantation <> SprevGroupPlantation then %>
							<td align="center">
								<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtMonth1=<%=v_txtMonth1%>">
								<%=rsgaji("GroupPlantation")%>
								</a>
							</td>
							<% SprevGroupPlantation = tempGroupPlantation %>	
						<% else	%>
							<td align="center"></td>					
						<% end if %>
					<% end if %>
					
					
					<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
						<%if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
							  tempGroupPlantation = rsgaji("GroupPlantation")
						   else
							  tempGroupPlantation = ""
						   end if
						
						   tempPlantation = rsgaji("Plantation")
						   tempGroupRegion = "--ALL--"
						   tempRegion = "--SUMMARY--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						%>
						<% if tempPlantation <> SprevPlantation then %>
							<td align="center">
								<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtMonth1=<%=v_txtMonth1%>">
								<%=rsgaji("Plantation")%>
								</a>
							</td>
							<% SprevPlantation = tempPlantation %>	
						<% else	%>
							<td align="center"></td>					
						<% end if %>	
					<% end if %>
					
					<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
						<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
							  tempGroupPlantation = rsgaji("GroupPlantation")
						   else
							  tempGroupPlantation = ""
						   end if					
						   
						   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
							  tempPlantation = rsgaji("Plantation")
						   else
							  tempPlantation = ""
						   end if
						   
						   tempGroupRegion = rsgaji("GroupRegion")
						   tempRegion = "--ALL--"
						   tempCompany ="--SUMMARY--"
						   tempfcba = "--SUMMARY--"
						%>
						<% if tempGroupRegion <> SprevGroupRegion then %>
							<td align="center">
								<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtMonth1=<%=v_txtMonth1%>">
								<%=rsgaji("GroupRegion")%>
								</a>
							</td>
							<% SprevGroupRegion= tempGroupRegion %>	
						<% else	%>
							<td align="center"></td>					
						<% end if %>	
					<% end if %>
					
					<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
						<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
							  tempGroupPlantation = rsgaji("GroupPlantation")
						   else
							  tempGroupPlantation = ""
						   end if					
						   
						   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
							  tempPlantation = rsgaji("Plantation")
						   else
							  tempPlantation = ""
						   end if
						   
						   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
							   tempGroupRegion = rsgaji("GroupRegion")
						   else
							   tempGroupRegion = ""
						   end if
						   
						   tempRegion = rsgaji("Region") 
						   tempCompany ="--ALL--"
						   tempfcba = "--SUMMARY--"
						%>
						
						<% if tempRegion <> SprevRegion then %>
							<td align="center">
								<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtMonth1=<%=v_txtMonth1%>">
								<%=rsgaji("region")%>
								</a>
							</td>
							<% SprevRegion = tempRegion %>	
						<% else	%>
							<td align="center"></td>					
						<% end if %>	
					<% end if %>
					
					<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
						<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
							  tempGroupPlantation = rsgaji("GroupPlantation")
						   else
							  tempGroupPlantation = ""
						   end if					
						   
						   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
							  tempPlantation = rsgaji("Plantation")
						   else
							  tempPlantation = ""
						   end if
						   
						   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
							   tempGroupRegion = rsgaji("GroupRegion")
						   else
							   tempGroupRegion = ""
						   end if
						   
						   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
							   tempRegion = rsgaji("Region") 
						   else
							   tempRegion = ""
						   end if 

						   tempCompany =rsgaji("Company")
						   tempfcba = "--ALL--"	   
						 %>
						
						<% if tempCompany <> SprevCompany then %>
							<td align="center">
								<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtMonth1=<%=v_txtMonth1%>">
								<%=rsgaji("company")%>
								</a>
							</td>
							<% SprevCompany = tempCompany %>	
						<% else	%>
							<td align="center"></td>					
						<% end if %>
					<% end if %>
					
					<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
						<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
							  tempGroupPlantation = rsgaji("GroupPlantation")
						   else
							  tempGroupPlantation = ""
						   end if					
						   
						   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
							  tempPlantation = rsgaji("Plantation")
						   else
							  tempPlantation = ""
						   end if
						   
						   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
							   tempGroupRegion = rsgaji("GroupRegion")
						   else
							   tempGroupRegion = ""
						   end if
						   
						   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
							   tempRegion = rsgaji("Region") 
						   else
							   tempRegion = ""
						   end if 
						   
						   if  v_txtcompany <>  "--SUMMARY--" and v_txtcompany <> "" then
							   tempCompany =rsgaji("Company")
						   else
							   tempCompany =""
						   end if

						   tempfcba = rsgaji("fcba")	   
						 %>
						
						<% if tempfcba <> sprevFcba then %>
							<td align="center">
								<a href="p_rekappayroll_2.asp?txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtMonth1=<%=v_txtMonth1%>">
								<%=rsgaji("fcba")%>
								</a>
							</td>
							<td align="left"><%=rsgaji("fcbaname")%></td>
							<% sprevFcba = tempfcba %>	
						<% else	%>
							<td align="center"></td>
							<td align="center"></td>
						<% end if %>
					<% end if %>

					<%If var_colspan = 1 Then%>
						<td align="right"></td>
					<%End If%>
						<td align="right"><%=FormatNumber(rsgaji("base_salary"), 0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsgaji("rapelgaji")), 0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsgaji("overtime")), 0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsgaji("premi")), 0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsgaji("tunjangan")), 0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsgaji("borongan")), 0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsgaji("p_jamsostek")), 0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsgaji("P_BPJSKES")), 0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsgaji("p_mk_p1")), 0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsgaji("p_panen")), 0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsgaji("p_pph21")), 0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsgaji("P_KOPERASI")), 0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsgaji("p_lain")), 0)%></td>
						<td align="right"><%=formatnumber( cdbl(rsgaji("total")), 0)%></td>
					<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
						<% currfcba = rsgaji("fcba") %>
						<% if currfcba <> prevfcba then %>
							<td align="right">
								<a href="P_RekapTransferGaji.asp?txtfcba=<%=currfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtMonth1=<%=v_txtMonth1%>">
								<%=formatnumber( cdbl(rsgaji("transfer")), 0)%>
								</a>
							</td>
							<% prevfcba = currfcba %>
						<% end if %>
					<%else%>
						<td align="right"><%=formatnumber( cdbl(rsgaji("transfer")), 0)%></td>
					<% end if %>
						<td align="right"><%=formatnumber( cdbl(rsgaji("nontransfer")), 0)%></td>
					</tr>
				<%End If%>
<%
				subtotalbase_salary = subtotalbase_salary + cdbl(rsgaji("base_salary"))
				subtotalrapelgaji = subtotalrapelgaji + Cdbl(rsgaji("rapelgaji"))
				subtotalovertime = subtotalovertime + Cdbl(rsgaji("overtime"))
				subtotalpremi = subtotalpremi + Cdbl(rsgaji("premi"))
				subtotaltunjangan = subtotaltunjangan + Cdbl(rsgaji("tunjangan"))
				subtotalborongan = subtotalborongan + Cdbl(rsgaji("borongan"))
				subtotalp_jamsostek = subtotalp_jamsostek + Cdbl(rsgaji("p_jamsostek"))
				subtotalp_bpjs = subtotalp_bpjs + Cdbl(rsgaji("P_BPJSKES"))
				subtotalp_mk_p1 = subtotalp_mk_p1 + Cdbl(rsgaji("p_mk_p1"))
				subtotalp_panen = subtotalp_panen + Cdbl(rsgaji("p_panen"))
				subtotalp_pph21 = subtotalp_pph21 + Cdbl(rsgaji("p_pph21"))
				subtotalp_koperasi = subtotalp_koperasi + Cdbl(rsgaji("P_KOPERASI"))
				subtotalp_lain = subtotalp_lain + Cdbl(rsgaji("p_lain"))
				subtotaltotal = subtotaltotal + Cdbl(rsgaji("total"))
				subtotaltransfer = subtotaltransfer + Cdbl(rsgaji("transfer"))
				subtotalnontransfer = subtotalnontransfer + Cdbl(rsgaji("nontransfer"))
%>
			<%Else%>
				<tr>
				
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% tempGroupPlantation = rsgaji("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtMonth1=<%=v_txtMonth1%>">
							<%=rsgaji("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<%if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsgaji("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					
					   tempPlantation = rsgaji("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtMonth1=<%=v_txtMonth1%>">
							<%=rsgaji("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsgaji("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsgaji("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsgaji("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtMonth1=<%=v_txtMonth1%>">
							<%=rsgaji("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsgaji("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsgaji("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsgaji("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsgaji("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtMonth1=<%=v_txtMonth1%>">
							<%=rsgaji("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsgaji("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsgaji("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsgaji("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsgaji("Region") 
					   else
					       tempRegion = ""
					   end if 

					   tempCompany =rsgaji("Company")
					   tempfcba = "--ALL--"	   
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtMonth1=<%=v_txtMonth1%>">
							<%=rsgaji("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsgaji("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsgaji("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsgaji("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsgaji("Region") 
					   else
					       tempRegion = ""
					   end if 
					   
					   if  v_txtcompany <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsgaji("Company")
					   else
					       tempCompany =""
					   end if

					   tempfcba = rsgaji("fcba")	   
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center">
							<a href="p_rekappayroll_2.asp?txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtMonth1=<%=v_txtMonth1%>">
						    <%=rsgaji("fcba")%>
							</a>
						</td>
						<td align="left"><%=rsgaji("fcbaname")%></td>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center"></td>
						<td align="center"></td>
					<% end if %>
				<% end if %>

				<%If var_colspan = 1 Then%>
					<td align="right"></td>
				<%End If%>
					<td align="right"><%=FormatNumber(rsgaji("base_salary"), 0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsgaji("rapelgaji")), 0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsgaji("overtime")), 0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsgaji("premi")), 0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsgaji("tunjangan")), 0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsgaji("borongan")), 0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsgaji("p_jamsostek")), 0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsgaji("P_BPJSKES")), 0)%></td>				
					<td align="right"><%=formatnumber( cdbl(rsgaji("p_mk_p1")), 0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsgaji("p_panen")), 0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsgaji("p_pph21")), 0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsgaji("P_KOPERASI")), 0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsgaji("p_lain")), 0)%></td>
					<td align="right"><%=formatnumber( cdbl(rsgaji("total")), 0)%></td>
					<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
						<% currfcba = rsgaji("fcba") %>
						<% if currfcba <> prevfcba then %>
							<td align="right">
								<a href="P_RekapTransferGaji.asp?txtfcba=<%=currfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtMonth1=<%=v_txtMonth1%>">
								<%=formatnumber( cdbl(rsgaji("transfer")), 0)%>
								</a>
							</td>
							<% prevfcba = currfcba %>
						<% end if %>
					<%else%>
						<td align="right"><%=formatnumber( cdbl(rsgaji("transfer")), 0)%></td>
					<% end if %>
						
					<td align="right"><%=formatnumber( cdbl(rsgaji("nontransfer")), 0)%></td>
				</tr>
			<%End If%>
<%
			grandtotalbase_salary = grandtotalbase_salary + Cdbl(rsgaji("base_salary"))
			grandtotalrapelgaji = grandtotalrapelgaji + Cdbl(rsgaji("rapelgaji"))
			grandtotalovertime = grandtotalovertime + Cdbl(rsgaji("overtime"))
			grandtotalpremi = grandtotalpremi + Cdbl(rsgaji("premi"))
			grandtotaltunjangan = grandtotaltunjangan + Cdbl(rsgaji("tunjangan"))
			grandtotalborongan = grandtotalborongan + Cdbl(rsgaji("borongan"))
			grandtotalp_jamsostek = grandtotalp_jamsostek + Cdbl(rsgaji("p_jamsostek"))
			grandtotalp_bpjs = grandtotalp_bpjs + Cdbl(rsgaji("P_BPJSKES"))
			grandtotalp_pph21 = grandtotalp_pph21 + Cdbl(rsgaji("p_pph21"))
			grandtotalp_mk_p1 = grandtotalp_mk_p1 + Cdbl(rsgaji("p_mk_p1"))
			grandtotalp_panen = grandtotalp_panen + Cdbl(rsgaji("p_panen"))
			grandtotalp_pph21 = grandtotalp_pph21 + Cdbl(rsgaji("p_pph21"))
			grandtotalp_koperasi = grandtotalp_koperasi + Cdbl(rsgaji("P_KOPERASI"))	
			grandtotalp_lain = grandtotalp_lain + Cdbl(rsgaji("p_lain"))
			grandtotaltotal = grandtotaltotal + Cdbl(rsgaji("total"))
			grandtotaltransfer = grandtotaltransfer + Cdbl(rsgaji("transfer"))
			grandtotalnontransfer = grandtotalnontransfer + Cdbl(rsgaji("nontransfer"))
			
			rsgaji.MoveNext
			PrevRegion = CurrRegion
		Wend	
	rsgaji.Close
	set rsgaji = Nothing
%>
			<%if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then%>
				<tr style="background-color:#F4F4A0;">
					<td align="left" colspan="<%=var_colspan-1%>" style="padding-left:10px;"><strong>Sub Total <%=PCase(PrevRegion)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(subtotalbase_salary, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(subtotalrapelgaji, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(subtotalovertime, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(subtotalpremi, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(subtotaltunjangan, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(subtotalborongan, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(subtotalp_jamsostek, 0)%></strong></td>
					
					<td align="Right"><strong><%=FormatNumber(subtotalp_mk_p1, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(subtotalp_panen, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(subtotalp_pph21, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(subtotalp_koperasi, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(subtotalp_lain, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(subtotaltotal, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(subtotaltransfer, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(subtotalnontransfer, 0)%></strong></td>
				</tr>
			<%End If%>
				<tr style="background-color:#F4F4A0;">
					<td align="left" colspan="<%=var_colspan-1%>" style="padding-left:10px;"><strong>Grand Total</strong></td>
					<td align="Right"><strong><%=FormatNumber(grandtotalbase_salary, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(grandtotalrapelgaji, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(grandtotalovertime, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(grandtotalpremi, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(grandtotaltunjangan, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(grandtotalborongan, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(grandtotalp_jamsostek, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(grandtotalp_bpjs, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(grandtotalp_mk_p1, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(grandtotalp_panen, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(grandtotalp_pph21, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(grandtotalp_koperasi, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(grandtotalp_lain, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(grandtotaltotal, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(grandtotaltransfer, 0)%></strong></td>
					<td align="Right"><strong><%=FormatNumber(grandtotalnontransfer, 0)%></strong></td>
				</tr>
<%End if%>
     </tbody>
</table>
	
   </body>
   </html>