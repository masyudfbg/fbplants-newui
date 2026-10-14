<%@Language=VBScript%>
<%
'Option Explicit
%>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtDateFrom, v_txtDateFrom2, v_txtDateTo2, v_txtDateTo,v_txtSupplier, CropName, FcbaType
Dim ScriptName, cClass, wherecons, selectCons, orderCons, q, CompanyName, fcbaname, MgrName, KTUName
Dim nBRUTO, nSORTASI, nNETTO, selectifkebun, selectifpabrik, v_txtfcba2
 

v_txtGroupPlantation = Request("txtGroupPlantation")
v_txtPlantation = Request("txtPlantation")
v_txtGroupRegion = Request("txtGroupRegion")
v_txtRegion = Request("txtRegion")
v_txtCompany = Request("txtCompany")
v_txtFcba = Request("txtFcba")
v_txtDateFrom = Request("txtDateFrom")

if "" & Request("txtDateFrom") = "" then v_txtDateFrom = textDate(date)
Dim arr_data, arr_eof, ii
 
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
Dim rsTerimaTBS, strTerimaTBS, nSpanTotal
' strTerimaTBS ="select Regioncode,CompanyCode,fcba,fieldcode,LongInterval,Afdeling,Harvestingsection,hectarageplanted " & _
              ' " from ( select businessunit_api.get_Groupplantation(fcba) Groupplantation, businessunit_api.get_plantation(fcba) " & _
			  ' "         PlantationCode, businessunit_api.get_Groupregion(fcba)     GroupRegionCode, businessunit_api.get_region(fcba) " & _
			  ' "         Regioncode, REGION_API.get_OrderData(businessunit_api.get_region(fcba)) orderregion, businessunit_api.get_CompanyCode(fcba) CompanyCode,company_api.get_OrderData(businessunit_api.get_CompanyCode(fcba)) ordercompany ,h.fcba , BUSINESSUNIT_API.get_OrderData(fcba) orderfcba, h.fieldcode, h.intervalnumber " & _
			  ' "         LongInterval, field_api.getValid_Afdeling(h.fieldcode,h.fcba,ondate) Afdeling," & _
              ' "			field_api.GetValid_Harvestingsection(h.fieldcode,h.fcba,ondate) Harvestingsection, " & _
			  ' "         field_api.GetValid_hectarageplanted(h.fieldcode,h.fcba,ondate)  hectarageplanted,  " & _ 
			  ' "         field_api.GetValid_Activation(h.fieldcode,h.fcba,ondate)  Activation,  " & _ 
			  ' "         field_api.GetValid_status(h.fieldcode,h.fcba,ondate)  status  " & _ 
			  ' "         from harvesting_interval_tab h where h.intervalnumber  > 10 and ondate = to_date('" & v_txtDateFrom & "', 'mm/dd/yyyy')  " & _
			  ' "      ) where  LongInterval = HARVESTING_INTERVAL_API.GetMaxPusinganPerDay(fcba,to_date('" & v_txtDateFrom & "', 'mm/dd/yyyy') ) " & _
			  ' "        AND status='TM'  and Activation='Y'" 


wherecons = " WHERE companycode IS NOT NULL "
			if v_txtGroupPlantation <> "--SUMMARY--"  and v_txtGroupPlantation <> "--ALL--" and v_txtGroupPlantation <> "" then
			   whereCons = wherecons & " And GroupPlantation = '" & v_txtGroupPlantation & "' "
			 
			end  if
			if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then
			    if v_txtPlantation <> "--ALL--" then
				   whereCons = whereCons & " And  PlantationCode = '" & v_txtPlantation & "' "
				end if   
				orderCons  = orderCons & ", PLANTATION_API.get_OrderData(Plantation) "
		        
			     
			end if
			if v_txtGroupRegion <> "--SUMMARY--"   and v_txtGroupRegion <> "" then
			    if v_txtGroupRegion <> "--ALL--" then
				   whereCons = whereCons & " And  GroupRegionCode = '" & v_txtGroupRegion & "' "
				end if
			end if
			if  v_txtRegion <> "--SUMMARY--"  and v_txtRegion <> "" then
			    if v_txtRegion <> "--ALL--" then
				   whereCons = whereCons & " And  Regioncode = '" & v_txtRegion & "' "
				end if
				
			end if
			if  v_txtCompany <> "--SUMMARY--" and v_txtCompany <> "" then
			    if v_txtCompany  <> "--ALL--" then
				   whereCons = whereCons & " And  companycode = '" & v_txtCompany & "' "
				end if
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
			    orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
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
					
'response.write strTerimaTBS & whereCons

strTerimaTBS = " SELECT regioncode," & _
               "        companycode," & _
               "        fcba," & _
               "        fieldcode," & _
               "        longinterval," & _
               "        afdeling," & _
               "        harvestingsection," & _
               "        hectarageplanted" & _
               " FROM (" & _
               " SELECT businessunit_api.get_groupplantation (aa.fcba) groupplantation," & _
               "        businessunit_api.get_plantation (aa.fcba) plantationcode," & _
               "        businessunit_api.get_groupregion (aa.fcba) groupregioncode," & _
               "        businessunit_api.get_region (aa.fcba) regioncode," & _
               "        region_api.get_orderdata (businessunit_api.get_region (aa.fcba)) orderregion," & _
               "        businessunit_api.get_companycode (aa.fcba) companycode," & _
               "        company_api.get_orderdata (businessunit_api.get_companycode (aa.fcba)) ordercompany," & _
               "        aa.fcba," & _
               "        businessunit_api.get_orderdata (aa.fcba) orderfcba," & _
               "        aa.fieldcode," & _
               "        aa.intervalnumber longinterval," & _
               "        field_api.getvalid_afdeling (aa.fieldcode, aa.fcba, aa.ondate) afdeling," & _
               "        field_api.getvalid_harvestingsection (aa.fieldcode, aa.fcba, aa.ondate) harvestingsection," & _
               "        field_api.getvalid_hectarageplanted (aa.fieldcode, aa.fcba, aa.ondate) hectarageplanted," & _
               "        field_api.getvalid_activation (aa.fieldcode, aa.fcba, aa.ondate) activation," & _
               "        field_api.getvalid_status (aa.fieldcode, aa.fcba, aa.ondate) status" & _
               " FROM harvesting_interval_tab aa" & _
               " INNER JOIN (" & _
               "     SELECT fcba, ondate, MAX(intervalnumber) intervalnumber" & _
               "     FROM harvesting_interval_tab" & _
               "     WHERE intervalnumber > 10" & _
               "           AND ondate = TO_DATE('" & v_txtDateFrom & "', 'mm/dd/yyyy')" & _
               "           AND field_api.getvalid_status(fieldcode, fcba, ondate) = 'TM'" & _
               "           AND field_api.getvalid_activation(fieldcode, fcba, ondate) = 'Y'" & _
			   "           AND field_api.getvalid_harvestingsection (fieldcode, fcba, ondate) IS NOT NULL" & _
               "     GROUP BY fcba, ondate" & _
               " ) hh" & _
               " ON (aa.fcba = hh.fcba AND aa.ondate = hh.ondate AND aa.intervalnumber = hh.intervalnumber)" & _
               " WHERE field_api.getvalid_status(aa.fieldcode, aa.fcba, aa.ondate) = 'TM'" & _
               "       AND field_api.getvalid_activation(aa.fieldcode, aa.fcba, aa.ondate) = 'Y'" & _
               "       " & _ 
               " )" & _
               "" & wherecons & "" & _
               " ORDER BY orderregion," & _
               "          ordercompany," & _
               "          orderfcba," & _
               "          Afdeling"
			   
arr_eof = true
Set rsTerimaTBS = Server.CreateObject("ADODB.Recordset")
rsTerimaTBS.Open strTerimaTBS, connstr, adOpenStatic
'--(Old) rsTerimaTBS.Open strTerimaTBS & whereCons & " order by orderregion,ordercompany ,orderfcba , Afdeling" , connstr, adOpenStatic

' response.write strTerimaTBS & whereCons & " order by orderregion,ordercompany ,orderfcba , Afdeling"
' response.write strTerimaTBS
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Pusingan Panen Terpanjang</title>
<link rel="stylesheet" href="/css/jquery-ui.css" />
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
    <script src="/js/jquery-1.9.1.js"></script>
	<script src="/js/jquery-ui.js"></script>
     <script>
	$(function() {
	$( ".datepicker" ).datepicker();
	});
	</script>
<script language="JavaScript" src="/js/date-picker.js"></script> 
<script language="JavaScript">
var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
</script>
</head>

<body>
	<div id="wrapper">
	<!--#include file="menutop.asp"-->
	<!--#include file="../print_excel.asp"-->
	<form action="<%=ScriptName%>" method="post" name="frmEntry">
	<div id="header-wrapper" style=" ">
	<!--#include virtual="/REPORT/FormSelectNewSumE.html"-->
    <table id="form-table" class="centered-margin">

	<tr>		
		<td><font face="verdana" size="1">Date</font></td>
		<td align="center">
			<input class="datepicker" type="text" maxlength="10" name="txtDateFrom" size="10" value="<%=v_txtDateFrom%>" title="Type Date mm/dd/yyyy" >
	    </td>
		<td><input type="submit" value="View" /></td>
	</tr>
	</table>
	</div>
	</form>

<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="50" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
		<h4>PT.<%=CompanyName%></h4>
		<h5><%=fcbaname%></h5>
	    </div>
		<div id="subheader" class="centered"  style="width:50%; margin:auto;">
		<h2>SUMMARY PUSINGAN PANEN TERPANJANG</h3>
		<br/>
			<h5>TANGGAL: <%=v_txtdatefrom%></h5>
		</div>
	</div>
 	</th>
   </tr><br>
	<tr>
	  <th rowspan="2" style="color:#FFFFFF" width='10%'>Group Kebun</th>
	  <th rowspan="2" style="color:#FFFFFF" width='10%'>PT</th>
	  <th rowspan="2" style="color:#FFFFFF" width='10%'>Kebun</th>
	  <th rowspan="2" style="color:#FFFFFF" width='15%'>Pusingan Terpanjang<br> Hari Ini</th>
	  <th colspan="4" style="color:#FFFFFF" width='45%'>Lokasi</th>
    </tr>
	<tr>
	    <th style="color:#FFFFFF" width='10%'>Blok</th>
		<th style="color:#FFFFFF" width='10%'>Afdeling</th>
		<th style="color:#FFFFFF" width='10%'>Seksi Panen</th>
		<th style="color:#FFFFFF" width='10%'>Luas</th>
	</tr>
    </thead>
	<tbody>
<%
Dim tempfcba,tempCompany,tempregion,templongintervalfcba,templongintervalCompany,templongintervalregion
Dim WEfcba,WECompany, WEregion

tempfcba    = ""
tempCompany = ""
tempregion  = ""
WEfcba      = 0
WECompany   = 0 
WEregion    = 0
templongintervalfcba = 0
templongintervalCompany = 0
templongintervalregion =0

Function CountColumnOccurances(curCol, curRow, arr)
  Dim occurances  : occurances = 1                  ' how many repeats
  Dim examinedRow : examinedRow = curRow + 1        ' the row we're comparing to
  Dim curValue    : curValue = arr(curCol, curRow)  ' the value were using to compare

  Do While examinedRow <= UBound(arr, 2)            ' the second dimension is row
    If arr(curCol, examinedRow) = curValue Then     ' the next row has the same value
      occurances = occurances + 1
    Else                                            ' the next row is different
      Exit Do
    End If
    examinedRow = examinedRow + 1
  Loop

  CountColumnOccurances = occurances
End Function

Function f_span_region(data_array__, region__)
	result_ = 0
	arr_1 = data_array__
	for indx = LBound(data_array__, 2) to UBound(data_array__, 2)
		if arr_1(0, indx) = region__ then
			result_ = result_ + 1
		end if
	next
	f_span_region = result_
End Function

Function f_span_company(data_array__, region__, company__)
	result_ = 0
	arr_1 = data_array__
	for indx = LBound(data_array__, 2) to UBound(data_array__, 2)
		if arr_1(0, indx) = region__ then
			if arr_1(1, indx) = company__ then
				result_ = result_ + 1
			end if
		end if
	next
	f_span_company = result_
End Function

Function f_count_block(data_array__, fcba__)
	result_ = 0
	arr_1 = data_array__
	for indx = LBound(data_array__, 2) to UBound(data_array__, 2)
		if arr_1(2, indx) = fcba__ then
				result_ = result_ + 1
		end if
	next
	f_count_block = result_
End Function

Function Get_ha(data_array, column_, wherecons)
	result_ = 0
	for indx = LBound(data_array, 2) to UBound(data_array, 2)
		if data_array(column_, indx) = wherecons then
			result_ = result_ + cdbl(data_array(7, indx))
		end if			
	next
	Get_ha = result_
End Function

Function Get_interval_ha(data_array, column_, wherecons)
	result_ = 0
	for indx = LBound(data_array, 2) to UBound(data_array, 2)
		if data_array(column_, indx) = wherecons then
			result_ = result_ + (cdbl(data_array(4, indx))*cdbl(data_array(7, indx)))
		end if			
	next
	Get_interval_ha = result_
End Function

Function ha_company(data_array, region__, company__)
	result_ = 0
	for indx = LBound(data_array, 2) to UBound(data_array, 2)
		if data_array(0, indx) = region__ then
			if data_array(1, indx) = company__ then
				result_ = result_ + cdbl(data_array(7, indx))
			end if
		end if
	next
	ha_company = result_
End Function

Function interval_ha_company(data_array, region__, company__)
	result_ = 0
	for indx = LBound(data_array, 2) to UBound(data_array, 2)
		if data_array(0, indx) = region__ then
			if data_array(1, indx) = company__ then
				result_ = result_ + (cdbl(data_array(4, indx))*cdbl(data_array(7, indx)))
			end if
		end if
	next
	interval_ha_company = result_
End Function

Function total_company_span(data_array__, region__, company__)
	total_fcba = 0
	total_blok = 0
	result__ = 0
	arr_1 = data_array__
	arr_2 = data_array__
	curr_B =""
	prev_B =""
	for nn = LBound(arr_1, 2) to UBound(arr_1, 2)
		if arr_1(0, nn) = region__ then
			if arr_1(1, nn) = company__ then
				curr_B = arr_1(2, nn)
				if curr_B <> prev_B then
					count_blok = 0
					for nnn = LBound(arr_2, 2) to UBound(arr_2, 2)
						if arr_2(0, nnn) = region__ and arr_2(1, nnn) = company__ and arr_2(2, nnn) = curr_B then
							count_blok = count_blok +1
						end if
					next

					if count_blok > 1 then
						result__ = result__ +1
					end if
				end if
			end if
		end if
		prev_B = curr_B
	next
	
	total_company_span = result__
End Function

Function total_region_span(data_array__, region__)
	total_fcba = 0
	result__ = 0
	arr_1 = data_array__
	arr_2 = data_array__
	arr_3 = data_array__
	
	curr_C =""
	prev_C =""
	
	curr_B =""
	prev_B =""
	for nn = LBound(arr_1, 2) to UBound(arr_1, 2)
		if arr_1(0, nn) = region__ then
			curr_C = arr_1(1, nn)
			if curr_C <> prev_C then
				total_fcba = 0
				for nnn = LBound(arr_2, 2) to UBound(arr_2, 2)
					if arr_2(0, nnn) = region__ and arr_2(1, nnn) = curr_C then
						curr_B = arr_2(2, nnn)
						if curr_B <> prev_B then
							total_fcba = total_fcba + 1
							count_blok = 0
							for nnnn = LBound(arr_3, 2) to UBound(arr_3, 2)
								if arr_3(0, nnnn) = region__ and arr_3(1, nnnn) = curr_C and arr_3(2, nnnn) = curr_B then
									count_blok = count_blok +1
								end if
							next

							if count_blok > 1 then
								result__ = result__ +1
							end if
							
							
						end if
					end if
					prev_B = curr_B
				next
				if total_fcba >= 1 then
					result__ = result__ +1
				end if
			end if			
		end if
		prev_C = curr_C
	next	
	
	total_region_span = result__
End Function


Function unit_of_company(data_array__, region__, company__)
	total_fcba = 0
	result__ = 0
	arr_1 = data_array__
	for nn = LBound(arr_1, 2) to UBound(arr_1, 2)
		if arr_1(0, nn) = region__ then
			if arr_1(1, nn) = company__ then
				curr_B = arr_1(2, nn)
				if curr_B <> prev_B then
					count_fcba = 1
					total_fcba = total_fcba + count_fcba
				end if
			end if
		end if
			prev_B = curr_B
	next
	result__ = total_fcba
	
	unit_of_company = result__
End Function

Function company_of_region(data_array__, region__)
	total_company = 0
	result__ = 0
	arr_1 = data_array__
	for nn = LBound(arr_1, 2) to UBound(arr_1, 2)
		if arr_1(0, nn) = region__ then
				curr_C = arr_1(1, nn)
				if curr_C <> prev_C then
					count_company = 1
					total_company = total_company + count_company
				end if
		end if
			prev_C = curr_C
	next
	result__ = total_company
	
	company_of_region = result__
End Function

if rsTerimaTBS.bof and rsTerimaTBS.eof then
%>
	<tr><td colspan="8">NO DATA AVAILABLE</td></tr>
<%else
	arr_data = rsTerimaTBS.getrows()
	arr_eof= false
end if
rsTerimaTBS.Close
set rsTerimaTBS = Nothing
%>
<%IF NOT arr_eof THEN%>
<%
RowSpan_R = 1
RowSpan_C = 1
RowSpan_B = 1
pr_R = 0
pr_C = 0
pr_B = 0

curr_R = ""
prev_R = ""
count_R = 0
For ii = LBound(arr_data, 2) To UBound(arr_data, 2)

	curr_R = arr_data(0, ii)
	if curr_R <> prev_R then
		count_R = count_R +1
	end if
	Response.Write "<tr>"
	pr_R = pr_R + 1
	pr_C = pr_C + 1
	pr_B = pr_B + 1
	if pr_R = 1 then
		post_RowSpan_R = 0
		'response.write total_region_span(arr_data, arr_data(0, ii))
		if total_region_span(arr_data, arr_data(0, ii)) >= 1 then
			post_RowSpan_R = total_region_span(arr_data, arr_data(0, ii))
		end if
		
		RowSpan_R = f_span_region(arr_data, arr_data(0, ii))
		response.write "<td rowspan='"&RowSpan_R+post_RowSpan_R&"'>" & arr_data(0, ii) & "</td>"
	end if
	if pr_C = 1 then
		post_RowSpan_C = 0
		if total_company_span(arr_data, arr_data(0, ii),  arr_data(1, ii)) > 0 then
			post_RowSpan_C = cdbl(total_company_span(arr_data, arr_data(0, ii), arr_data(1, ii)))
		end if
		
		RowSpan_C = f_span_company(arr_data, arr_data(0, ii), arr_data(1, ii))
		response.write "<td rowspan='"&RowSpan_C+post_RowSpan_C&"'>" & arr_data(1, ii) & "</td>"
	end if
	if pr_B = 1 then
		RowSpan_B = CountColumnOccurances(2, ii, arr_data)
		response.write "<td rowspan='"&RowSpan_B&"'>" & arr_data(2, ii) & "</td>"
	end if
%>
			<td align="right"><%=FormatNumber(arr_data(4, ii),0)%></td>
			<td align="center"><%=arr_data(3, ii)%></td>
			<td align="center"><%=arr_data(5, ii)%></td>
			<td align="center"><%=arr_data(6, ii)%></td>
			<td align="right"><%=arr_data(7, ii)%></td>
	
<%
	if pr_B = RowSpan_B then
		pr_B = 0
		jumlah_blok = f_count_block(arr_data, arr_data(2, ii))
		ha_B = Get_ha(arr_data, 2, arr_data(2, ii))
		interval_ha_B = Get_interval_ha(arr_data, 2, arr_data(2, ii))
		
		if ha_B = 0 then
			rata2_pusingan = 0
		else
			rata2_pusingan = Round(cdbl(interval_ha_B)/cdbl(ha_B), 0)
		end if
		if jumlah_blok > 1 then
			response.write "</tr>"
			response.write "<tr style='background-color:#92d050;'>" & _
			               "    <td><b>SUB TOTAL "&arr_data(2, ii)&"</b></td>" & _
						   "    <td align='right'><b>"&cdbl(rata2_pusingan)&"</b></td><td colspan='3'>&nbsp;</td>" & _
						   "    <td align='right'><b>"&cdbl(ha_B)&"</b></td>" & _
						   "</tr>"
		end if
		
	end if
	if pr_C = RowSpan_C then
		pr_C = 0
		if unit_of_company(arr_data, arr_data(0, ii),  arr_data(1, ii)) >= 1 then
			ha_B = ha_company(arr_data, arr_data(0, ii), arr_data(1, ii))
			interval_ha_B = interval_ha_company(arr_data, arr_data(0, ii), arr_data(1, ii))
			
			if ha_B = 0 then
				rata2_pusingan = 0
			else
				rata2_pusingan = Round(cdbl(interval_ha_B)/cdbl(ha_B), 0)
			end if
			response.write "</tr>"
			response.write "<tr style='background-color:#F4F4A0;'>" & _
						   "    <td colspan='2'><b>TOTAL "&arr_data(1, ii)&"</b></td>" & _
						   "    <td align='right'><b>"&cdbl(rata2_pusingan)&"</b></td><td colspan='3'>&nbsp;</td>" & _
						   "    <td align='right'><b>"&cdbl(ha_B)&"</b></td>" & _
						   "</tr>"
		end if
	end if
	if pr_R = RowSpan_R then
		pr_R = 0
		if company_of_region(arr_data, arr_data(0, ii)) >= 1 then
			ha_B = Get_ha(arr_data, 0, arr_data(0, ii))
			interval_ha_B = Get_interval_ha(arr_data, 0, arr_data(0, ii))
			
			if ha_B = 0 then
				rata2_pusingan = 0
			else
				rata2_pusingan = Round(cdbl(interval_ha_B)/cdbl(ha_B), 0)
			end if
			response.write "</tr>"
			response.write "<tr style='background-color:#ffff00;'>" & _
						   "    <td colspan='3'><b>TOTAL "&arr_data(0, ii)&"</b></td>" & _
						   "    <td align='right'><b>"&cdbl(rata2_pusingan)&"</b></td><td colspan='3'>&nbsp;</td>" & _
						   "    <td align='right'><b>"&cdbl(ha_B)&"</b></td>" & _
						   "</tr>"
		end if
	end if
	
	prev_R = curr_R
Next

if count_R > 1 then
gt_ha_region = 0
gt_interval_ha_region = 0
gt_rata2_pusingan = 0
for xx = LBound(arr_data, 2) To UBound(arr_data, 2)
	gt_ha_region = gt_ha_region + cdbl(arr_data(7, xx))
	gt_interval_ha_region = gt_interval_ha_region + (cdbl(arr_data(4, xx))*cdbl(arr_data(7, xx)))
next

if gt_ha_region = 0 then
	gt_rata2_pusingan = 0
else
	gt_rata2_pusingan = Round(cdbl(gt_interval_ha_region)/cdbl(gt_ha_region), 0)
end if
response.write "<tr><td colspan='8'></td></tr><tr style='background-color:#ffff00;'>" & _
			   "    <td colspan='3'><b>GRAND TOTAL</b></td>" & _
			   "    <td align='right'><b>"&cdbl(gt_rata2_pusingan)&"</b></td><td colspan='3'>&nbsp;</td>" & _
			   "    <td align='right'><b>"&cdbl(gt_ha_region)&"</b></td>" & _
			   "</tr>"
end if
%>
<%END IF%>
</tbody>
</table>
</body>
</html>