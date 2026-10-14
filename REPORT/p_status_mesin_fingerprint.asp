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
d_span = 35

Dim CompanyName, q, fcbaname, MgrName, KTUName, CropName, FcbaType

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

function f_fcbaspan (arr__, where__)
	result__ = 0
	arr1 = arr__
	
	for xyy=0 to ubound(arr1,2)
		if cstr("" & arr1(1, xyy)) = cstr("" & where__) then
				result__ = result__ +1
		end if
	next
	
	f_fcbaspan = result__
end function

nSpanTotal = 9

v_Isactive = ""
v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")
v_Isactive              = Request("txtIsactive")
txtstatus = Request("txtstatus")

If v_Isactive = "" then v_Isactive = "TRUE"
If txtstatus = "" then txtstatus = "TRUE"


txtDate = request("txtDate")

if ""&txtDate = "" then txtDate = textDate(date)

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
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " "
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
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " "
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
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " "
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
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " "
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

If Len(Trim(v_txtfcba)) = 0 Then v_txtfcba = "--SUMMARY--"

If Len(Trim(v_txtfcba)) > 0 Then
	if v_txtfcba = "--SUMMARY--" then
		selectCons	=	selectCons	& " , fcba"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , nourut"
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtfcba = "--ALL--" then
		selectCons	=	selectCons	& " , fcba"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , nourut"
	Else
		selectCons	=	selectCons	& " , fcba"
		whereCons	=	whereCons	& " And fcba = '" & v_txtfcba & "'"
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

If v_Isactive <> "ALL" then
   whereCons	=	whereCons	& " and businessunit_api.Isactive(fcba) = '" & v_Isactive & "'"	
end if

txtYear = Year(txtDate)
txtMonth = Month(txtDate)
txttdate = Request("txttdate")
txtjabatan = Request("txtjabatan")

If "" & txtYear = "" Then txtYear = Year(txtDate)
If "" & txtMonth = "" Then txtMonth = Month(txtDate)
If "" & txttdate = "" Then txttdate = textDate(date)

txttdate = txtMonth & "/01/" & txtYear

MaxDate = Day(txtDate)


ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain = " SELECT" & strFieldCons & _
          "        , serial_number" & _
          "        , location" & _
          "        , locationname" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 1, total, NULL)) day_1" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 2, total, NULL)) day_2" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 3, total, NULL)) day_3" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 4, total, NULL)) day_4" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 5, total, NULL)) day_5" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 6, total, NULL)) day_6" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 7, total, NULL)) day_7" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 8, total, NULL)) day_8" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 9, total, NULL)) day_9" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 10, total, NULL)) day_10" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 11, total, NULL)) day_11" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 12, total, NULL)) day_12" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 13, total, NULL)) day_13" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 14, total, NULL)) day_14" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 15, total, NULL)) day_15" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 16, total, NULL)) day_16" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 17, total, NULL)) day_17" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 18, total, NULL)) day_18" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 19, total, NULL)) day_19" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 20, total, NULL)) day_20" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 21, total, NULL)) day_21" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 22, total, NULL)) day_22" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 23, total, NULL)) day_23" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 24, total, NULL)) day_24" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 25, total, NULL)) day_25" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 26, total, NULL)) day_26" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 27, total, NULL)) day_27" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 28, total, NULL)) day_28" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 29, total, NULL)) day_29" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 30, total, NULL)) day_30" & _
          "        , SUM(DECODE(TO_NUMBER(TO_CHAR(tdate, 'dd')), 31, total, NULL)) day_31" & _
          " FROM (" & _
          "     SELECT master.fcba, master.serial_number, master.location, master.locationname, master.tdate, NVL(data.total,0) total" & _
          "     FROM (" & _
          "         SELECT *" & _
          "         FROM master_machine_fingerprint mf" & _
          "              , ( SELECT TO_DATE ('"&txttdate&"', 'MM/DD/YYYY') +ROWNUM -1 tdate" & _
          "                  FROM all_objects" & _
          "                  WHERE ROWNUM <= LAST_DAY(TO_DATE ('"&txttdate&"', 'MM/DD/YYYY')) - TO_DATE ('"&txttdate&"', 'MM/DD/YYYY') + 1) dd" & _
          "         WHERE activation = '" & txtstatus & "'" & _
          "     ) master" & _
          "     LEFT JOIN (" & _
          "         SELECT REGEXP_REPLACE(machine_sn, '[^0-9]', '') serial_number, TRUNC(checktime) tdate, COUNT(*) total" & _
          "         FROM userfinger.checkinout" & _
          "         WHERE TRUNC(checktime) BETWEEN TO_DATE ('"&txttdate&"', 'MM/DD/YYYY') AND LAST_DAY(TO_DATE ('"&txttdate&"', 'MM/DD/YYYY'))" & _
          "         GROUP BY machine_sn, TRUNC(checktime)" & _
          "     ) data" & _
          "     ON (master.serial_number = data.serial_number AND master.tdate = data.tdate)" & _
          " )" & _
          " GROUP BY fcba" & _
          "          , serial_number" & _
          "          , location" & _
          "          , locationname"
								
strmonitoringclosing2 = " , serial_number, location, locationname" & _
                        " , day_1, day_2, day_3, day_4, day_5, day_6, day_7, day_8, day_9, day_10" & _
                        " , day_11, day_12, day_13, day_14, day_15, day_16, day_17, day_18, day_19, day_20" & _
						" , day_21, day_22, day_23, day_24, day_25, day_26, day_27, day_28, day_29, day_30, day_31"

strmonitoringclosing = " SELECT GroupPlantation" & selectCons & strmonitoringclosing2 &"  FROM (" & strMain & " )" & whereCons & _
                       " ORDER BY nourut, locationname"

'Response.write strmonitoringclosing

arr_field_eof = true
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open strmonitoringclosing, connstr, adOpenStatic

If Not qORS.EOF Then
	arr_field = qORS.getrows()
	arr_field_eof = false
End If
qORS.Close
set qORS = Nothing
%>
<html>

<head>
<title>ePlaS Reporting - Monitoring Closing</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<link rel="stylesheet" href="/css/font/font-awesome.css" />
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
.circle {
	width: 10px;
	height: 10px;
	-moz-border-radius: 50px;
	-webkit-border-radius: 50px;
	border-radius: 50px;
}
.green{
	color: #2ECC40;
}
.red{
	color: #FF4136;
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
							<td><font face="verdana" size="1">Tanggal</font></td>
							<td width="2">
								<input class="datepicker" type="text" name="txtDate" value="<%=txtDate%>" />
							</td>
							<td><font face="verdana" size="1">Status</font></td>
							<td>
								<select name="txtstatus">
									<option value="TRUE" <%if cstr(""&txtstatus) = "TRUE" then%> selected <%end if%>>TRUE</option>
									<option value="FALSE" <%if cstr(""&txtstatus) = "FALSE" then%> selected <%end if%>>FALSE</option>
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

<table border="0" width="100%" id="product-table" class="xthf-site" >
	<thead style="border-left:1px #000000 solid;">
	<tr>
         <th colspan="<%=d_span +add_colspan%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
		 <div id="header-wrapper">
			<div id="head_judul">
				<h4><%If Len(Trim(varCompanyName)) > 0 Then%>PT. <%End If%><%=varCompanyName%></h4>
				<h5><%=varFcbaName%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>STATUS MESIN FINGER</h2>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(txtMonth)%>&nbsp;<%=txtYear%></h4>
			</div>
          </div>  
		 </th>
	</tr>
	<tr>
		<th rowspan="2"><%=titleSite%></th>
		<th rowspan="2">Serial Number</th>
		<th rowspan="2">Location</th>
	</tr>
	<tr>
		<%for xi=1 to cint(MaxDate)%>
		<th style="min-width:20px;max-width:20px;"><%=xi%></th>
		<%next%>
	</tr>

    </thead>
	<tbody>
<%
if not arr_field_eof then
	currfcba = ""
	prevfcba = ""
	for y=0 to ubound(arr_field,2)
			response.write "<tr>"
			currfcba = arr_field(1,y)
			if currfcba <> prevfcba then			
				response.write "	<td rowspan='"&f_fcbaspan(arr_field, arr_field(1,y))&"'align='center' valign='top'>"&arr_field(1,y)&"</td>"
			end if
			response.write "	<td align='left'>"&arr_field(2,y)&"</td>"
			response.write "	<td align='left'>"&arr_field(4,y)&"</td>"
			for xi=5 to (5 + (Cint(MaxDate) -1) )
				if cdbl(to_nol(arr_field(xi,y))) = 0 then
					'response.write "	<td align='center'><i class='fa fa-times red'></i></td>"
					'response.write "	<td align='center'><i class='fa red'>&#xf00d;</i></td>"
					response.write "	<td align='center'><img src='/images/mark-x.png' /></td>"
				else
					'response.write "	<td align='center'><i class='fa fa-check'></i></td>"
					'response.write "	<td align='center'><i class='fa'>&#xf00c;</i></td>"
					response.write "	<td align='center'><img src='/images/mark-check.png' /></td>"
				end if
			next
			response.write "</tr>"
			
			prevfcba = currfcba
	next
else
	response.write "<tr><td colspan='"&d_span&"'>NO DATA AVAILABLE</td></tr>"
end if
%>
	</tbody>
</table>

   </body>
   </html>