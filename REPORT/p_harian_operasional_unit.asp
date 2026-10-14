<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtYear, v_txtMonth, v_txtJobcode, v_txtAfdeling, v_txtTahunTanam, v_txtBlok, v_Isactive
Dim ScriptName
Dim rsmonitoringclosing, strmonitoringclosing, strMain, strmonitoringclosing2
Dim NO, cClass, nday, xNum, nSpanTotal

Dim var_colspan : var_colspan = 1

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType

vehiclegroup = request("vehiclegroup")
kelas = request("kelas")
txtDate = request("txtDate")
if "" & txtDate = "" then txtDate = textDate(date)

function f_color (datefrom__, dateto__)
	result__ = "background-color:#f53333;"
	if "" & datefrom__ <> "" and "" & dateto__ <> "" then
		if DateDiff("d",Cdate(datefrom__),CDate(dateto__)) >= 0 and DateDiff("d",Cdate(datefrom__),CDate(dateto__)) <= 2 then
			result__ = "background-color:#92d050;"
		elseif DateDiff("d",Cdate(datefrom__),CDate(dateto__)) > 2 then
			result__ = "background-color:#ffc000;"
		else
			result__ = "background-color:#f53333;"
		end if
	end if
	f_color = result__
end function

function f_nool (amount__)
	result__ = "-"
	if cdbl(to_nol(amount__)) <> 0 then
		result__ = cdbl(to_nol(amount__))
	else
		result__ = "-"
	end if
	f_nool = result__
end function

nSpanTotal = 23

v_Isactive = ""
v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")
v_Isactive              = Request("txtIsactive")

If v_Isactive = "" then
   v_Isactive="TRUE"
end if




Dim selectFields, strFieldCons, selectCons, whereCons, GroupbyCons, orderCons

strFieldCons	=	" "
					
wherecons	=	" "

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
		whereCons	=	wherecons	& " AND bu.GroupPlantation = '" & v_txtGroupPlantation & "'"
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
		whereCons	=	whereCons	& " AND bu.Plantation = '" & v_txtPlantation & "'"
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
		whereCons	=	whereCons	& " AND bu.GroupRegion = '" & v_txtGroupRegion & "'"
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
		whereCons	=	whereCons	& " AND bu.Region = '" & v_txtRegion & "'"
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
		whereCons	=	whereCons	& " AND bu.fccompanycode = '" & v_txtCompany & "'"
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

If Len(Trim(v_txtfcba)) = 0 Then v_txtfcba = "--SUMMARY--"

If Len(Trim(v_txtfcba)) > 0 Then
	if v_txtfcba = "--SUMMARY--" then
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , nourut"
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtfcba = "--ALL--" then
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , nourut"
	Else
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	whereCons	& " AND vo.fcba = '" & v_txtfcba & "'"
		orderCons	=	orderCons	& " , nourut"
		
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


v_txtYear = Request("txtYear")
v_txtMonth = Request("txtMonth")

If "" & v_txtYear = "" Then v_txtYear = Year(Session("v_Costbook"))
If "" & v_txtMonth = "" Then v_txtMonth = Month(Session("v_Costbook"))


ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

if "" & vehiclegroup <> "" then whereCons = whereCons & " AND vh.vehiclegroupcode = '" & vehiclegroup & "'"
if "" & kelas <> "" then whereCons = whereCons & " AND vh.vehiclesubgroupcode = '" & kelas & "'"


str_hi = " SELECT vo.fcba, vo.vehiclecode, vo.remarks, vo.vehiclestatus, NVL(unit, 0) unit_hi, NVL(fuel, 0) fuel_hi, NVL(harihujan, 0) harihujan_hi, NVL(curahhujan, 0) curahhujan_hi" & _
         " FROM vehicle vh" & _
         " LEFT JOIN vehicleoperation vo ON (vo.fcba = vh.fcba AND vo.vehiclecode = vh.fccode)" & _
         " LEFT JOIN businessunit bu ON vh.fcba = bu.fccode" & _
		 " WHERE vo.vehicledate = TO_DATE('" & txtDate & "', 'mm/dd/yyyy')" & _
		 "       AND USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "', vo.fcba) IS NOT NULL" & _
		 "       AND bu.isactive = 'TRUE'" & _
		 "       " & whereCons
		 
str_shi = " SELECT vo.fcba, vo.vehiclecode, NVL(SUM(unit), 0) unit_shi, NVL(SUM(fuel), 0) fuel_shi, NVL(SUM(harihujan), 0) harihujan_shi, NVL(SUM(curahhujan), 0) curahhujan_shi" & _
          " FROM vehicle vh" & _
          " LEFT JOIN vehicleoperation vo ON (vo.fcba = vh.fcba AND vo.vehiclecode = vh.fccode)" & _
          " LEFT JOIN businessunit bu ON vh.fcba = bu.fccode" & _
		  " WHERE TO_CHAR(vo.vehicledate, 'yyyymm') = TO_CHAR(TO_DATE('" & txtDate & "', 'mm/dd/yyyy'), 'yyyymm')" & _
		  "       AND vo.vehicledate <= TO_DATE('" & txtDate & "', 'mm/dd/yyyy')" & _
		  "       AND USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "', vo.fcba) IS NOT NULL" & _
		  "       AND bu.isactive = 'TRUE'" & _
		  "       " & whereCons & "" & _
		  " GROUP BY vo.fcba, vo.vehiclecode"
		  
strJoin = " SELECT b.fcba, b.vehiclecode, a.remarks, a.vehiclestatus, NVL(a.unit_hi, 0) unit_hi, NVL(b.unit_shi, 0) unit_shi, NVL(a.fuel_hi, 0) fuel_hi, NVL(b.fuel_shi, 0) fuel_shi, NVL(a.harihujan_hi, 0) harihujan_hi, NVL(b.harihujan_shi, 0) harihujan_shi, NVL(a.curahhujan_hi, 0) curahhujan_hi, NVL(b.curahhujan_shi, 0) curahhujan_shi" & _
          " FROM ("&str_shi&") b" & _
		  " LEFT JOIN ("&str_hi&") a ON (a.fcba = b.fcba AND a.vehiclecode = b.vehiclecode)"

strMain		=	" SELECT gr.fcname groupregion" & _
				"        , r.fcname region" & _
				"        , bu.fccompanycode" & _
				"        , vo.fcba" & _
				"        , vh.vehiclegroupcode" & _
				"        , vo.vehiclecode" & _
				"        , vh.make" & _
				"        , vh.model" & _
				"        , vsg.description vehiclesubgroupname" & _
				"        , TO_CHAR(vh.yearofmade, 'yyyy') yearofmade" & _
				"        , 1 totalofunit" & _
				"        , CASE WHEN vo.vehiclestatus = '1' THEN 1 ELSE 0 END baik" & _
				"        , CASE WHEN vo.vehiclestatus = '2' THEN 1 ELSE 0 END standby" & _
				"        , CASE WHEN vo.vehiclestatus = '3' THEN 1 ELSE 0 END rusak" & _
				"        , CASE WHEN vo.vehiclestatus = '4' THEN 1 ELSE 0 END writeoff" & _
				"        , vo.unit_hi" & _
				"        , vo.unit_shi" & _
				"        , vo.fuel_hi" & _
				"        , vo.fuel_shi" & _
				"        , vo.harihujan_hi" & _
				"        , vo.harihujan_shi" & _
				"        , vo.curahhujan_hi" & _
				"        , vo.curahhujan_shi" & _
				"        , vo.remarks" & _
				"        , gr.orderdata order_gr, r.orderdata order_r, bu.orderdata order_bu" & _
				" FROM ("&strJoin&") vo" & _
				" INNER JOIN vehicle vh ON (vo.fcba = vh.fcba AND vo.vehiclecode = vh.fccode)" & _
				" INNER JOIN vehiclegroup vhg ON (vh.fcba = vhg.fcba AND vh.vehiclegroupcode = vhg.vehiclegroupcode)" & _
				" INNER JOIN vehiclesubgroup vsg ON (vh.fcba = vsg.fcba AND vh.vehiclesubgroupcode = vsg.vehiclesubgroupcode)" & _
				" INNER JOIN businessunit bu ON vo.fcba = bu.fccode" & _
				" INNER JOIN region r ON bu.region = r.fccode" & _
				" INNER JOIN groupregion gr ON bu.groupregion = gr.fccode"
								
strmonitoringclosing2 = " "

strmonitoringclosing = " SELECT *" & _
                       " FROM (" & strMain & ") " & _
                       " ORDER BY order_gr,  order_r,  order_bu, vehiclegroupcode, vehiclecode"

'Response.write strmonitoringclosing
%>
<html>

<head>
<title>ePlaS Reporting - Harian Operasional Unit/ Alat</title>
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
						<td><font face="verdana" size="1">Vehicle Group</font></td>
						<td>
							<select name="vehiclegroup">
								<option value="" <% If "" & vehiclegroup = "" Then %> Selected <%End if%>>--All--</option>
								<%
								q = " SELECT vehiclegroupcode fccode, description fcname FROM vehiclegroup where fcba = 'KSB'"
								Set rsComboA = Server.CreateObject("ADODB.Recordset")
								
								rsComboA.Open q, connstr
								if not rsComboA.eof then
									Do While Not rsComboA.eof
								%>
									<option value="<%=rsComboA("fccode")%>" <% If vehiclegroup = rsComboA("fccode") Then %> Selected <%End if%>><%=rsComboA("fcname")%></option> 
								<% 
										rsComboA.movenext
									loop
								End if
								rsComboA.Close : set rsComboA = Nothing : q = ""
								%>
							</select>
						</td>
						<td><font face="verdana" size="1">Kelas</font></td>
						<td>
							<select name="kelas">
								<option value="" <% If "" & kelas = "" Then %> Selected <%End if%>>--All--</option>
								<%
								q = " SELECT vehiclesubgroupcode fccode, description fcname FROM vehiclesubgroup where fcba = 'KSB'"
								Set rsComboA = Server.CreateObject("ADODB.Recordset")
								
								rsComboA.Open q, connstr
								if not rsComboA.eof then
									Do While Not rsComboA.eof
								%>
									<option value="<%=rsComboA("fccode")%>" <% If kelas = rsComboA("fccode") Then %> Selected <%End if%>><%=rsComboA("fcname")%></option> 
								<% 
										rsComboA.movenext
									loop
								End if
								rsComboA.Close : set rsComboA = Nothing : q = ""
								%>
							</select>
						</td>
						<td><font face="verdana" size="1">Tanggal</font></td>
						<td>
							<input class="datepicker" type="text" name="txtDate" value="<%=txtDate%>" size="15">
						</td>
						<td>
							<input type="submit" value="View" />
						</td>
					</tr>
				</table>
			</div>
			</form>
		</div>

<table width="100%" id="product-table" class="xthf-site" style="border-bottom:1px #3A5F0B solid;">
	<thead style="border-left:1px #000000 solid;">
		<tr>
			<th colspan="26" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
				<div id="header-wrapper">
		<div id="head_judul">
				<% if  CompanyName <>  "" then %>
					<h4>PT.<%=CompanyName%></h4> 
				<% End if%>
				<% If  FcbaName <> ""  then %>
					   <h5><%=FcbaName%></h5>   
				<% End if%>
				</div>
				<div id="subheader" class="centered">
					<h2>LAPORAN HARIAN OPERASIONAL UNIT/ALAT</h2>
					<br/>
					<h4>Tanggal : <%=Day(txtDate) & " " & Bulan(Month(txtDate)) & " " & Year(txtDate)%></h4>
				</div>
              </div
			</th>
		</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;" rowspan="2"><%=titleGroupRegion%></th>
		<th style="background-color:#00A300;color:#fff;" rowspan="2"><%=titleRegion%></th>
		<th style="background-color:#00A300;color:#fff;" rowspan="2"><%=titleCompany%></th>
		<th style="background-color:#00A300;color:#fff;" rowspan="2"><%=titleSite%></th>		
		<th style="background-color:#00A300;color:#fff;" rowspan="2">Vehicle Group</th>
		<th style="background-color:#00A300;color:#fff;" rowspan="2">Vehicle Code</th>
		<th style="background-color:#00A300;color:#fff;" rowspan="2">Merk</th>
		<th style="background-color:#00A300;color:#fff;" rowspan="2">Type</th>
		<th style="background-color:#00A300;color:#fff;" rowspan="2">Kelas</th>
		<th style="background-color:#00A300;color:#fff;" rowspan="2">Tahun</th>
		<th style="background-color:#00A300;color:#fff;" rowspan="2">Total Unit</th>
		<th style="background-color:#00A300;color:#fff;" colspan="2">Availability</th>
		<th style="background-color:#00A300;color:#fff;" colspan="2">Break Down</th>
		<th style="background-color:#00A300;color:#fff;" colspan="2">Hm-Km-Jam</th>
		<th style="background-color:#00A300;color:#fff;" colspan="2">Fuel (Ltr)</th>
		<th style="background-color:#00A300;color:#fff;" colspan="2">Ratio (Ltr/Hm-Km-Jam)</th>
		<th style="background-color:#00A300;color:#fff;" colspan="2">Hari Hujan (Hari)</th>
		<th style="background-color:#00A300;color:#fff;" colspan="2">Curah Hujan (mm)</th>
		<th style="background-color:#00A300;color:#fff;" rowspan="2">Keterangan</th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;">Baik</th>
		<th style="background-color:#00A300;color:#fff;">Stand By</th>
		<th style="background-color:#00A300;color:#fff;">Rusak</th>
		<th style="background-color:#00A300;color:#fff;">W/ Off</th>
		<th style="background-color:#00A300;color:#fff;">HI</th>
		<th style="background-color:#00A300;color:#fff;">SHI</th>
		<th style="background-color:#00A300;color:#fff;">HI</th>
		<th style="background-color:#00A300;color:#fff;">SHI</th>
		<th style="background-color:#00A300;color:#fff;">HI</th>
		<th style="background-color:#00A300;color:#fff;">SHI</th>
		<th style="background-color:#00A300;color:#fff;">HI</th>
		<th style="background-color:#00A300;color:#fff;">SHI</th>
		<th style="background-color:#00A300;color:#fff;">HI</th>
		<th style="background-color:#00A300;color:#fff;">SHI</th>
	</tr>

    </thead>
	<tbody>
<%
Set rsmonitoringclosing = Server.CreateObject("ADODB.Recordset")
rsmonitoringclosing.Open strmonitoringclosing, connstr, adOpenStatic

Dim Rp_Per_Ha_RowReal : Rp_Per_Ha_RowReal = 0
Dim Rp_Per_Ha_RowBudget : Rp_Per_Ha_RowBudget = 0
Dim RowPercentage : RowPercentage = 0

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


if rsmonitoringclosing.bof and rsmonitoringclosing.eof then
%>
	<tr>
		<td colspan="26">NO DATA AVAILABLE</td>
	</tr>
<%
else
		while not rsmonitoringclosing.eof
			if cdbl(to_nol(rsmonitoringclosing("unit_hi"))) = 0 then
				ratio_hi = 0
			else
				ratio_hi = Round( cdbl(to_nol(rsmonitoringclosing("fuel_hi")))/cdbl(to_nol(rsmonitoringclosing("unit_hi"))), 2 )
			end if
			
			if cdbl(to_nol(rsmonitoringclosing("unit_shi"))) = 0 then
				ratio_shi = 0
			else
				ratio_shi = Round( cdbl(to_nol(rsmonitoringclosing("fuel_shi")))/cdbl(to_nol(rsmonitoringclosing("unit_shi"))), 2 )
			end if
%>
				<tr>
                    <td align="left"><%=rsmonitoringclosing("groupregion")%></td>  
					<td align="left"><%=rsmonitoringclosing("region")%></td>
					<td align="left"><%=rsmonitoringclosing("fccompanycode")%></td>
					<td align="center"><%=rsmonitoringclosing("fcba")%></td>
					<td align="center"><%=rsmonitoringclosing("vehiclegroupcode")%></td>
					<td align="left"><%=rsmonitoringclosing("vehiclecode")%></td>
					<td align="left"><%=rsmonitoringclosing("make")%></td>
					<td align="left"><%=rsmonitoringclosing("model")%></td>
					<td align="left"><%=rsmonitoringclosing("vehiclesubgroupname")%></td>
					<td align="center"><%=rsmonitoringclosing("yearofmade")%></td>
					<td align="right"><%=rsmonitoringclosing("totalofunit")%></td>
					<td align="right"><%=f_nool(rsmonitoringclosing("baik"))%></td>
					<td align="right"><%=f_nool(rsmonitoringclosing("standby"))%></td>
					<td align="right"><%=f_nool(rsmonitoringclosing("rusak"))%></td>
					<td align="right"><%=f_nool(rsmonitoringclosing("writeoff"))%></td>
					<td align="right"><%=f_nool(rsmonitoringclosing("unit_hi"))%></td>
					<td align="right"><%=f_nool(rsmonitoringclosing("unit_shi"))%></td>
					<td align="right"><%=f_nool(rsmonitoringclosing("fuel_hi"))%></td>
					<td align="right"><%=f_nool(rsmonitoringclosing("fuel_shi"))%></td>
					<td align="right"><%=f_nool(ratio_hi)%></td>
					<td align="right"><%=f_nool(ratio_shi)%></td>
					<td align="right"><%=f_nool(rsmonitoringclosing("harihujan_hi"))%></td>
					<td align="right"><%=f_nool(rsmonitoringclosing("harihujan_shi"))%></td>
					<td align="right"><%=f_nool(rsmonitoringclosing("curahhujan_hi"))%></td>
					<td align="right"><%=f_nool(rsmonitoringclosing("curahhujan_shi"))%></td>
					<td align="left"><%=rsmonitoringclosing("remarks")%></td>
				</tr>
<%
			rsmonitoringclosing.MoveNext
			Wend
rsmonitoringclosing.Close
set rsmonitoringclosing = Nothing
%>
<%End if%>
	</tbody>
</table>

   </body>
   </html>