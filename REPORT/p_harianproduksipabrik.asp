<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtYear, v_txtMonth, v_txtCustomer,v_txtOwnstate,v_txtstorage_code,v_txtproduct_code
Dim ScriptName
Dim rspabrik, strprodukpabrik, strMain, strsumkg,no
Dim cClass, nday, xNum

Dim var_ColSpan : var_ColSpan = 0
Dim jml_showdata : jml_showdata = 7

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType,v_txtCUSTFFBTYPE,conCUSTFFBTYPE,conCUSTFFBTYPE2

v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")
v_txtCUSTFFBTYPE		= Request("CUSTFFBTYPE")
v_txtOwnstate			= Request("txtOwnstate")
v_txtstorage_code		= Request("txtStorage_code")
v_txtproduct_code		= Request("txtProduct_code")


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
		var_ColSpan	=	var_ColSpan	+ 1
	Else
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " And GroupPlantation = '" & v_txtGroupPlantation & "'"
		orderCons	=	orderCons	& " "
		var_ColSpan	=	var_ColSpan	+ 1
	End If
End If

If Len(Trim(v_txtPlantation)) > 0 Then
	if v_txtPlantation = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtPlantation = "--ALL--" then
		selectCons	=	selectCons	& " , Plantation"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_ColSpan	=	var_ColSpan	+ 1
	Else
		selectCons	=	selectCons	& " , Plantation"
		whereCons	=	whereCons	& " And Plantation = '" & v_txtPlantation & "'"
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_ColSpan	=	var_ColSpan	+ 1
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
		var_ColSpan	=	var_ColSpan	+ 1
	Else
		selectCons	=	selectCons	& " , GroupRegion"
		whereCons	=	whereCons	& " And GroupRegion = '" & v_txtGroupRegion & "'"
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_ColSpan	=	var_ColSpan	+ 1
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
		var_ColSpan	=	var_ColSpan	+ 1
	Else
		selectCons	=	selectCons	& " , Region"
		whereCons	=	whereCons	& " And Region = '" & v_txtRegion & "'"
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_ColSpan	=	var_ColSpan	+ 1
	End If
End If


If Len(Trim(v_txtCompany)) > 0 Then
	if v_txtCompany = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtCompany = "--ALL--" then
		selectCons	=	selectCons	& " , company"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		var_ColSpan	=	var_ColSpan	+ 1
	Else
		selectCons	=	selectCons	& " , company"
		whereCons	=	whereCons	& " And company = '" & v_txtCompany & "'"
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		var_ColSpan	=	var_ColSpan	+ 1
		
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
		var_ColSpan	=	var_ColSpan	+ 1
	Else
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	whereCons	& " And fcba = '" & v_txtfcba & "'"
		orderCons	=	orderCons	& " , nourut"
		var_ColSpan	=	var_ColSpan	+ 1
		
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

If Len(Trim(v_txtproduct_code)) = 0 Then
	v_txtproduct_code	=	"KER"
End If
	selectCons	=	selectCons	& " , product_code"
	orderCons	=	orderCons	& " , product_code"
	var_ColSpan	=	var_ColSpan	+ 1
	
If Len(Trim(v_txtstorage_code)) = 0 Then
	v_txtstorage_code	=	"--SUMMARY--"
End If
	If v_txtstorage_code = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
	Elseif v_txtstorage_code = "--ALL--" then
		selectCons	=	selectCons	& " , storage_code"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , storage_code"
		var_ColSpan	=	var_ColSpan	+ 1
	
	End if
	


ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

	Dim strloop_st,strloop_sy,strloop_slt,strloop_qty,strloop_qtyused,strloop_pr
	for nday = 1 to 31
		strloop_st = strloop_st + ", SUM(DECODE(TO_NUMBER (TO_CHAR (TDATE, 'DD')), " & nday & ", sounding_today, 0)) KG_" & nday
		strloop_sy = strloop_sy + ", SUM(DECODE(TO_NUMBER (TO_CHAR (TDATE, 'DD')), " & nday & ", sounding_yesterday, 0)) SY_" & nday
		strloop_slt = strloop_slt + ", SUM(DECODE(TO_NUMBER (TO_CHAR (TDATE, 'DD')), " & nday & ", sales_today, 0)) SLT_" & nday
		strloop_qty = strloop_qty + ", SUM(DECODE(TO_NUMBER (TO_CHAR (TDATE, 'DD')), " & nday & ", qty_transfer, 0)) QTY_" & nday
		strloop_qty = strloop_qty + ", SUM(DECODE(TO_NUMBER (TO_CHAR (TDATE, 'DD')), " & nday & ", qty_receive, 0)) QTYRCV_" & nday
		strloop_qtyused = strloop_qtyused + ", SUM(DECODE(TO_NUMBER (TO_CHAR (TDATE, 'DD')), " & nday & ", qty_used, 0)) QTYUSED_" & nday
		strloop_pr = strloop_pr + ", SUM(DECODE(TO_NUMBER (TO_CHAR (TDATE, 'DD')), " & nday & ", production_today, 0)) PR_" & nday
	next
	
strMain		=	" SELECT" & strFieldCons & _ 
				" , product_code" & _
				" , storage_code" & _
				strloop_st & _
				strloop_sy & _
				strloop_slt & _
				strloop_qty & _
				strloop_qtyused & _
				strloop_pr & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (TDATE, 'YYYY')), " & v_txtYear & ", sounding_today, TO_NUMBER (TO_CHAR (TDATE, 'MM')), " & v_txtMonth & ", sounding_today, 0)) KG_TOTAL" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (TDATE, 'YYYY')), " & v_txtYear & ", sounding_yesterday, TO_NUMBER (TO_CHAR (TDATE, 'MM')), " & v_txtMonth & ", sounding_yesterday, 0)) SY_TOTAL" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (TDATE, 'YYYY')), " & v_txtYear & ", sales_today, TO_NUMBER (TO_CHAR (TDATE, 'MM')), " & v_txtMonth & ", sales_today, 0)) SLT_TOTAL" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (TDATE, 'YYYY')), " & v_txtYear & ", qty_transfer, TO_NUMBER (TO_CHAR (TDATE, 'MM')), " & v_txtMonth & ", qty_transfer, 0)) QTY_TOTAL" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (TDATE, 'YYYY')), " & v_txtYear & ", qty_receive, TO_NUMBER (TO_CHAR (TDATE, 'MM')), " & v_txtMonth & ", qty_receive, 0)) QTYRCV_TOTAL" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (TDATE, 'YYYY')), " & v_txtYear & ", qty_used, TO_NUMBER (TO_CHAR (TDATE, 'MM')), " & v_txtMonth & ", qty_used, 0)) QTYUSED_TOTAL" & _
				"     , SUM(DECODE(TO_NUMBER (TO_CHAR (TDATE, 'YYYY')), " & v_txtYear & ", production_today, TO_NUMBER (TO_CHAR (TDATE, 'MM')), " & v_txtMonth & ", production_today, 0)) PR_TOTAL" & _
				" FROM (" & _
				"     SELECT" & _
				"         FCBA" & _
				"         , product_code" & _
				"         , storage_code" & _
				"         , TDATE" & _
				"         , sum(sounding_today) sounding_today" & _
				"         , sum(sounding_yesterday) sounding_yesterday" & _
				"         , sum(sales_today) sales_today" & _
				"         , sum(qty_transfer) qty_transfer" & _
				"         , sum(qty_receive) qty_receive" & _
				"         , sum(qty_used) qty_used" & _
				"         , sum(production_today) production_today" & _
				"     FROM (" & _
				"         SELECT" & _
				"             FCBA" & _
				"             , product_code" & _
				"             , storage_code" & _
				"             , TDATE" & _
				"             , sounding_today" & _
				"             , sounding_yesterday" & _
				"             , sales_today" & _
				"             , qty_transfer" & _
				"             , qty_receive" & _
				"             , qty_used" & _
				"             , production_today" & _
				"         FROM" & _
				"             mill_production mp" & _
				"         WHERE" & _
				"             product_code = '" & v_txtproduct_code & "'" & _
				"             AND TO_NUMBER (TO_CHAR (TDATE, 'YYYY')) = " & v_txtYear & "" & _
				"             AND TO_NUMBER (TO_CHAR (TDATE, 'MM')) = " & v_txtMonth & "" & _
				"         )" & _
				"     GROUP BY" & _
				"         FCBA" & _
				"         , product_code" & _
				"         , storage_code" & _
				"         , TDATE" & _
				"         , sounding_today" & _
				"         , sounding_yesterday" & _
				"         , sales_today" & _
				"         , qty_transfer" & _
				"         , production_today" & _
				"     ) WHERE USER_COMPANY_API.GetfcbaType('" & Session("CodeName") & "',fcba) <> 'E'" & _
				" GROUP BY" & _
				"     FCBA" & _
				"     , product_code" & _
				"     , storage_code" & _
				" ORDER BY" & _
				"     FCBA"								
Dim sum_st,sum_sy,sum_slt,sum_qty,sum_qtyrcv,sum_qtyused,sum_pr	
	for nday = 1 to 31
		sum_st = sum_st + ", SUM(KG_" & nday & ") KG_" & nday
		sum_sy = sum_sy + ", SUM(SY_" & nday & ") SY_" & nday
		sum_slt = sum_slt + ", SUM(SLT_" & nday & ") SLT_" & nday
		sum_qty = sum_qty + ", SUM(QTY_" & nday & ") QTY_" & nday
		sum_qtyrcv = sum_qtyrcv + ", SUM(QTYRCV_" & nday & ") QTYRCV_" & nday
		sum_qtyused = sum_qtyused + ", SUM(QTYUSED_" & nday & ") QTYUSED_" & nday
		sum_pr = sum_pr + ", SUM(PR_" & nday & ") PR_" & nday
	next
strsumkg	=	sum_st & sum_sy & sum_slt & sum_qty & sum_qtyrcv & sum_qtyused & sum_pr & _
				"     , SUM(KG_TOTAL) KG_TOTAL" & _
				"     , SUM(SY_TOTAL) SY_TOTAL" & _
				"     , SUM(SLT_TOTAL) SLT_TOTAL" & _
				"     , SUM(QTY_TOTAL) QTY_TOTAL" & _
				"     , SUM(QTYRCV_TOTAL) QTYRCV_TOTAL" & _
				"     , SUM(QTYUSED_TOTAL) QTYUSED_TOTAL" & _
				"     , SUM(PR_TOTAL) PR_TOTAL"

strprodukpabrik	=	" select GroupPlantation " & selectCons & strsumkg & " From (" & strMain & ")" & _
				whereCons & _
				" group by GroupPlantation" & selectCons & " " & _
				" Order by GroupPlantation" & orderCons 

' response.write strprodukpabrik

%>
<html>
<head>
<title>ePlaS Reporting - Laporan Monitoring Mutasi Stock</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<!--<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
  }, false
);
</script>-->
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
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="menutop.asp"-->
			</div>
			<form action="<%=ScriptName%>" method="post" name="frmEntry">
			<div class="printhide">
				<!--#include virtual="/REPORT/FormSelectNewSumM.html"-->
				
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
							<font face="verdana" size="1">Product Code</font>
						</td>
						<td>
						
							<select name="txtProduct_code">
								<option VALUE="KER" <% if v_txtproduct_code = "KER" then %>SELECTED<% end if %>>KER</option>
								<option VALUE="CPO" <% if v_txtproduct_code = "CPO" then %>SELECTED<% end if %>>CPO</option>
								<option VALUE="CPKO" <% if v_txtproduct_code = "CPKO" then %>SELECTED<% end if %>>CPKO</option>
								<option VALUE="PKE" <% if v_txtproduct_code = "PKE" then %>SELECTED<% end if %>>PKE</option>
								<option VALUE="SHL" <% if v_txtproduct_code = "SHL" then %>SELECTED<% end if %>>SHELL</option>
								
							</select>
						</td>	
						<td>
							<font face="verdana" size="1">Storage Code</font>
						</td>
						<td>
						
							<select name="txtStorage_code">
								<option VALUE="--SUMMARY--" <% if v_txtstorage_code = "--SUMMARY--" then %>SELECTED<% end if %>>--SUMMARY--</option>
								<option VALUE="--ALL--" <% if v_txtstorage_code = "--ALL--" then %>SELECTED<% end if %>>--ALL--</option>
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
			<th colspan="<%=var_ColSpan+(jml_showdata+7)%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
				<div id="header">
				<% if  CompanyName <>  "" then %>
					<h4>PT.<%=CompanyName%></h4> 
				<% End if%>
				<% If  FcbaName <> ""  then %>
					<h5><%=FcbaName%></h5>   
				<% End if%>
				</div>
				<div id="subheader" class="centered">
					<h2>LAPORAN MONITORING MUTASI STOCK</h3>
					<br/>
					<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
				</div>
			</th>
		</tr>
		<tr>

		<%If Len(Trim(v_txtGroupPlantation)) > 0 And v_txtGroupPlantation <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleGroupPlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtPlantation)) > 0 And v_txtPlantation <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titlePlantation%></th>
		<%End If%>

		<%If Len(Trim(v_txtGroupRegion)) > 0 And v_txtGroupRegion <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleGroupRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtRegion)) > 0 And v_txtRegion <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleRegion%></th>
		<%End If%>

		<%If Len(Trim(v_txtCompany)) > 0 And v_txtCompany <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleCompany%></th>
		<%End If%>

		<%If Len(Trim(v_txtfcba)) > 0 And v_txtfcba <>  "--SUMMARY--" Then%>
			<th rowspan="3" style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
		<%End If%>
		

			<th rowspan="3" style="background-color:#00A300;color:#fff;">Product Code</th>

		
		<% IF v_txtstorage_code <> "--SUMMARY--"  then %>
			<th rowspan="3" style="background-color:#00A300;color:#fff;">Storage Code</th>
		<% END IF %>

			<th colspan="<%=jml_showdata*31%>" style="background-color:#00A300;color:#fff;">TANGGAL</th>
			<th rowspan="2" colspan="<%=jml_showdata%>" style="background-color:#00A300;color:#fff;">TOTAL</th>
		</tr>
		<tr>
			<%for nday = 1 to 31%>
			<th colspan="7" style="background-color:#00A300;color:#fff;"><%=nday%></th>
			<%next%>
		</tr>
		<tr>
			<%for nday = 1 to 32%>
				<th style="background-color:#00A300;color:#fff;min-width:70px;">Snd Tday</th>
				<th style="background-color:#00A300;color:#fff;min-width:70px;">Snd Yday</th>
				<th style="background-color:#00A300;color:#fff;min-width:70px;">Sls Tday</th>
				<th style="background-color:#00A300;color:#fff;min-width:70px;">Qty Trans</th>
				<th style="background-color:#00A300;color:#fff;min-width:70px;">Qty Receive</th>
				<th style="background-color:#00A300;color:#fff;min-width:70px;">Qty Used</th>
				<th style="background-color:#00A300;color:#fff;min-width:70px;">Prod Tday</th>
			<%next%>
		</tr>
</thead>
<tbody>
<%
'response.write strprodukpabrik

Dim GrandSLT(31)
Dim GrandQTY(31)
Dim GrandQTYRCV(31)
Dim GrandQTYUSED(31)
Dim GrandPR(31)
Dim GrandTOTAL(7)
Dim grandday

For grandday = 1 to 31
     GrandSLT(grandday)  = 0
	 GrandQTY(grandday)  = 0
	 GrandQTYRCV(grandday)  = 0
	 GrandQTYUSED(grandday)  = 0
	 GrandPR(grandday)  = 0
next

For grandday = 1 to 7
     GrandTOTAL(grandday)  = 0
next

Dim GrandKg_Total : GrandKg_Total = 0
'response.write strprodukpabrik
Set rspabrik = Server.CreateObject("ADODB.Recordset")
rspabrik.Open strprodukpabrik, connstr, adOpenStatic

Dim tempGroupPlantation : tempGroupPlantation = ""
Dim tempPlantation : tempPlantation = ""
Dim tempGroupRegion : tempGroupRegion = ""
Dim tempRegion : tempRegion = ""
Dim tempCompany : tempCompany = ""
Dim tempfcba : tempfcba = ""
Dim tempCustomer : tempCustomer = ""

Dim SprevGroupPlantation : SprevGroupPlantation = ""
Dim SprevPlantation : SprevPlantation = ""
Dim SprevGroupRegion : SprevGroupRegion = ""
Dim SprevRegion : SprevRegion = ""
Dim SprevCompany : SprevCompany = ""
Dim sprevFcba : sprevFcba = ""
Dim sPrevCustomer : sPrevCustomer = ""

if rspabrik.bof and rspabrik.eof then
%>
	<tr>
		<td colspan="<%=var_ColSpan+(jml_showdata*32)%>">NO DATA AVAILABLE</td>
	</tr>
<%
else
		while not rspabrik.eof
%>
				<tr>
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% tempGroupPlantation = rspabrik("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtCustomer=<%=tempCustomer%>">
							<%=rspabrik("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rspabrik("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					   tempPlantation = rspabrik("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtCustomer=<%=tempCustomer%>">
							<%=rspabrik("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rspabrik("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rspabrik("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rspabrik("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtCustomer=<%=tempCustomer%>">
							<%=rspabrik("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <%  if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rspabrik("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rspabrik("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--"  and  v_txtGroupRegion <> "" then
					       tempGroupRegion = rspabrik("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rspabrik("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtCustomer=<%=tempCustomer%>">
							<%=rspabrik("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% 
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rspabrik("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rspabrik("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rspabrik("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rspabrik("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rspabrik("Company")
					   tempfcba = "--ALL--"			   
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtCustomer=<%=tempCustomer%>">
							<%=rspabrik("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rspabrik("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rspabrik("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion  <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rspabrik("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion  <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rspabrik("Region") 
					   else
					       tempRegion = ""
					   end if
					   
					   if  v_txtcompany  <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rspabrik("Company")
					   else
					       tempCompany =""
					   end if
					   
					   tempfcba = rspabrik("fcba")
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtYear=<%=v_txtYear%>&txtMonth=<%=v_txtMonth%>&txtCustomer=<%=tempCustomer%>">
						    <%=rspabrik("fcba")%>
							</a>
						</td>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
					
					<td align="center"><%=rspabrik("product_code")%></td>
				
				<% IF v_txtstorage_code <> "--SUMMARY--"  then %>
				   <td align="center"><%=rspabrik("storage_code")%></td>
				<% END IF %>
		
				<%For nday = 1 to 31%>
					<td align="right"><%=FormatNumber(rspabrik("KG_" & nday & ""), 2)%></td>
					<td align="right"><%=FormatNumber(rspabrik("SY_" & nday & ""), 2)%></td>
					<td align="right">

						<a class="td_link" href="../p_DELIVERYNOTE.asp?postingdate=<%=v_txtMonth & "/" & nday & "/" & v_txtYear  %>" target="_blank" >
							<%=FormatNumber(rspabrik("SLT_" & nday & ""), 0)%>
						</a>
					</td>
					<td align="right"><%=FormatNumber(rspabrik("QTY_" & nday & ""), 2)%></td>
					<td align="right"><%=FormatNumber(rspabrik("QTYRCV_" & nday & ""), 2)%></td>
					<td align="right"><%=FormatNumber(rspabrik("QTYUSED_" & nday & ""), 2)%></td>
					<td align="right"><%=FormatNumber(rspabrik("PR_" & nday & ""), 2)%></td>
					<%
                     GrandSLT(nday) =  GrandSLT(nday) + cCur(rspabrik("SLT_" & nday & ""))
					 GrandQTY(nday) =  GrandQTY(nday) + cCur(rspabrik("QTY_" & nday & ""))
					 GrandQTYRCV(nday) =  GrandQTYRCV(nday) + cCur(rspabrik("QTYRCV_" & nday & ""))
					 GrandQTYUSED(nday) =  GrandQTYUSED(nday) + cCur(rspabrik("QTYUSED_" & nday & ""))
					 GrandPR(nday)  =  GrandPR(nday) + cCur(rspabrik("PR_" & nday & ""))
			
					%>
				<%Next%>
				
					<td align="right"><%=FormatNumber(rspabrik("KG_TOTAL"), 2)%></td>
					<td align="right"><%=FormatNumber(rspabrik("SY_TOTAL"), 2)%></td>
					<td align="right"><%=FormatNumber(rspabrik("SLT_TOTAL"), 2)%></td>
					<td align="right"><%=FormatNumber(rspabrik("QTY_TOTAL"), 2)%></td>
					<td align="right"><%=FormatNumber(rspabrik("QTYRCV_TOTAL"), 2)%></td>
					<td align="right"><%=FormatNumber(rspabrik("QTYUSED_TOTAL"), 2)%></td>
					<td align="right"><%=FormatNumber(rspabrik("PR_TOTAL"), 2)%></td>
				</tr>
				<%
                     GrandTOTAL(1) =  GrandTOTAL(1) + cCur(rspabrik("KG_TOTAL"))
					 GrandTOTAL(2) =  GrandTOTAL(2) + cCur(rspabrik("SY_TOTAL"))
					 GrandTOTAL(3) =  GrandTOTAL(3) + cCur(rspabrik("SLT_TOTAL"))
					 GrandTOTAL(4) =  GrandTOTAL(4) + cCur(rspabrik("QTY_TOTAL"))
					 GrandTOTAL(5) =  GrandTOTAL(5) + cCur(rspabrik("QTYRCV_TOTAL"))
					 GrandTOTAL(6) =  GrandTOTAL(6) + cCur(rspabrik("QTYUSED_TOTAL"))
					 GrandTOTAL(7) =  GrandTOTAL(7) + cCur(rspabrik("PR_TOTAL"))
			
			rspabrik.MoveNext
			
			Wend	
rspabrik.Close
set rspabrik = Nothing
%>
	<tr bgcolor="#F4F4A0">
		<td colspan="<%=var_colspan%>" align="center"><strong>GRAND TOTAL</strong></td>	
		<% For nday = 1 to 31 %>	
			      <td colspan="2" align="right"><strong></strong></td>
		      <% If GrandSLT(nday)= 0 Then%>
			      <td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		      <%Else%>
			      <td align="right"><strong><%=FormatNumber(GrandSLT(nday), 2)%></strong></td>
		      <%End If%>
			  
			  <% If GrandQTY(nday)= 0 Then%>
			      <td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		      <%Else%>
			      <td align="right"><strong><%=FormatNumber(GrandQTY(nday), 2)%></strong></td>
		      <%End If%>
			  
			  <% If GrandQTYRCV(nday)= 0 Then%>
			      <td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		      <%Else%>
			      <td align="right"><strong><%=FormatNumber(GrandQTYRCV(nday), 2)%></strong></td>
		      <%End If%>
			  
			   <% If GrandQTYUSED(nday)= 0 Then%>
			      <td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		      <%Else%>
			      <td align="right"><strong><%=FormatNumber(GrandQTYUSED(nday), 2)%></strong></td>
		      <%End If%>
			  
			   <% If GrandPR(nday)= 0 Then%>
			      <td align="right"><strong><%=FormatNumber(0, 2)%></strong></td>
		      <%Else%>
			      <td align="right"><strong><%=FormatNumber(GrandPR(nday), 2)%></strong></td>
		      <%End If
		  next%>
		<td align="right"><strong><%=FormatNumber(GrandTOTAL(1), 2)%></strong></td>  
		<td align="right"><strong><%=FormatNumber(GrandTOTAL(2), 2)%></strong></td>  
		<td align="right"><strong><%=FormatNumber(GrandTOTAL(3), 2)%></strong></td>  
		<td align="right"><strong><%=FormatNumber(GrandTOTAL(4), 2)%></strong></td>  
		<td align="right"><strong><%=FormatNumber(GrandTOTAL(5), 2)%></strong></td>  
		<td align="right"><strong><%=FormatNumber(GrandTOTAL(6), 2)%></strong></td>  
		<td align="right"><strong><%=FormatNumber(GrandTOTAL(7), 2)%></strong></td>  
	</tr>
<%End if%>
	</tbody>
</table>

   </body>
</html>
