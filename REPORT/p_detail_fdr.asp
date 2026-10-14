<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/conndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtFcba, v_txtYear, v_txtMonth,v_txtDay,v_txtDate, v_txtAfdeling, v_txtBlok, v_txtColor
Dim ScriptName
Dim rsPusingan, strPusingan, strsql
Dim NO, cClass, nday,ColorID
Dim rsDefaultBa, strDefaultBa, varDefaultBa
Dim rsBaInfo, strBaInfo, varFcbaName,varCompanyCode, varCompanyName
Dim oRS, oSQL
maxCol = 0

txtYear = Request("txtYear")
txtMonth = Request("txtMonth")
txtindeks = Request("txtindeks")

If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))
If "" & txtMonth = "" Then txtMonth = Month(Session("v_Costbook"))
If "" & txtindeks = "" Then txtindeks = "01"

add_col_bud_prod_qty = ""
d_span = 14
	

v_txtAfdeling	=	Request("txtAfdeling")
v_txtBlok		=	Request("txtBlok")
v_txtFcba		=	Request("txtFcba")

arr_indeks_eof = true
oSql = " SELECT fccode, fcname" & _
       " FROM setup_bisa_pps_tab" & _
	   " WHERE levelid = '1'"
set oRs = server.CreateObject("Adodb.recordset")
oRs.Open oSql, connstr
if Not oRs.Eof then
	arr_indeks = oRs.getrows()
	arr_indeks_eof = false
End if
oRs.Close
Set oRs = Nothing

Function f_header(data_array__, where__)
	result_ = ""
	arr_1 = data_array__
	
	total__ = 0
	running__ = 0
	runningx__ = 1
	aaa = Split(where__,",")
	for each xxx in aaa
		total__ = total__ +1
	next
	
	for indx = 0 to UBound(arr_1, 2)
		if INSTR( cstr("" & where__), cstr("" & arr_1(0, indx)) ) <> 0 then
			running__ = running__ +1
			runningx__ = runningx__ +1
			result_ = result_ & cstr("" & arr_1(1, indx))
			if total__ > 1 then
				if running__ <> total__ then
					if runningx__ = total__ then
						result_ = result_ & " dan "
					else
						result_ = result_ & ", "
					end if
				end if
			end if
		end if
	next
	f_header = result_
End Function


strDefaultBa = " SELECT fccode AS fcba" & _
				" FROM businessunit bu" & _
				" INNER JOIN user_fcba_tab uft ON (bu.fccode = uft.fcba) " & _
				" WHERE fctype = 'E'" & _
				"     AND fcusercode = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	listba=0
	do while not rsDefaultBa.EOF
		listba = listba +1
		if listba = 1 then varDefaultBa = rsDefaultBa("FCBA")
		rsDefaultBa.movenext
	loop
Else
	varDefaultBa = ""
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

If "" & v_txtFcba = "" Then v_txtFcba = varDefaultBa

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

max_bisa = 0
strPusingan	=	" SELECT NVL(COUNT(*),0) total FROM setup_bisa_pps_tab WHERE levelid = '1' AND INSTR('"&txtindeks&"', fccode) <> 0"
'response.write strPusingan
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open strPusingan, connstr, adOpenStatic

If Not qORS.EOF Then
	max_bisa = qORS("total")
else
	max_bisa = 0
End If
qORS.Close
set qORS = Nothing

arr_field_eof = true
strPusingan	= " SELECT fcba" & _
              "        , ondate" & _
              "        , ombrometer" & _
              "        , rhp" & _
              "        , rh" & _
              "        , rainfall" & _
              "        , rainlessday" & _
              "        , totalrainfallinfifteendays" & _
              "        , fuelcondition" & _
              "        , rd" & _
              "        , trl" & _
              "        , fc" & _
              "        , fdi" & _
              "        , fdr_api.fdr(fdi) fdr" & _
              "        , fdr_api.fdr_color(fdi) fdr_color" & _
              " FROM (" & _
              "     SELECT fcba" & _
              "            , ondate" & _
              "            , ombrometer" & _
              "            , rhp" & _
              "            , rh" & _
              "            , rainfall" & _
              "            , rainlessday" & _
              "            , totalrainfallinfifteendays" & _
              "            , fuelcondition" & _
              "            , rd" & _
              "            , trl" & _
              "            , fc" & _
              "            , fdr_api.fdi(rh, rd, trl, fc) fdi" & _
              "     FROM (" & _
              "         SELECT fcba" & _
              "                , ondate" & _
              "                , ombrometer" & _
              "                , 50 rhp" & _
              "                , 100 rh" & _
              "                , rainfall" & _
              "                , rainlessday" & _
              "                , totalrainfallinfifteendays" & _
              "                , fdr_api.fuelcondition(rainlessday) fuelcondition" & _
              "                , fdr_api.rd(rainlessday) rd" & _
              "                , fdr_api.trl(totalrainfallinfifteendays) trl" & _
              "                , fdr_api.fc(fdr_api.fuelcondition(rainlessday)) fc" & _
              "         FROM fdr_tab" & _
              "         WHERE TO_CHAR(ondate, 'yyyymm') = '" & txtyear & "'||LPAD('" & txtmonth & "',2,0)" & _
              "               AND fcba = '" & v_txtFcba & "'"
if ""& v_txtAfdeling <> "" then strPusingan = strPusingan & "               AND ombrometer = '" & v_txtAfdeling & "'"
strPusingan = strPusingan & "     )" & _
              " ) ORDER BY ombrometer, ondate"
			  
'response.write strPusingan
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open strPusingan, connstr, adOpenStatic

If Not qORS.EOF Then
	arr_field = qORS.getrows()
	arr_field_eof = false
End If
qORS.Close
set qORS = Nothing


Function f_subtotal(data_array__, fccode__, cols__)
	result_ = 0
	tresult_ = 0
	arr_1 = data_array__
	
	for indx = 0 to UBound(arr_1, 2)
		if cdbl(to_nol(arr_1(2, indx))) = 1 then
			if left(cstr("" & arr_1(0, indx)), 4) = left(cstr("" & fccode__), 4) then
				result_ = result_ + cdbl(to_nol(arr_1(cols__, indx)))
			end if
		end if
	next
	
	for indx = 0 to UBound(arr_1, 2)
		if cdbl(to_nol(arr_1(2, indx))) = 1 then
			if left(cstr("" & arr_1(0, indx)), 4) = left(cstr("" & fccode__), 4) then
				tresult_ = tresult_ +1
			end if
		end if
	next
	
	if tresult_ = 0 then fresult_ = 0 else fresult_ = round(result_/tresult_*100,0)
	f_subtotal = fresult_
End Function


strBaInfo = " SELECT" & _
			"     FCNAME as FCBANAME" & _
			"    , fccompanycode  , FCCOMPANYNAME as COMPANYNAME" & _
			" FROM BUSINESSUNIT BU" & _
			" WHERE" & _
			"     FCCODE = '" & v_txtFcba & "'"

Set rsBaInfo = Server.CreateObject("ADODB.Recordset")
rsBaInfo.Open strBaInfo, connstr, adOpenStatic

If Not rsBaInfo.EOF Then
	varFcbaName = rsBaInfo("FCBANAME")
	varCompanyName = rsBaInfo("COMPANYNAME")
	varCompanyCode = rsBaInfo("fccompanycode")
Else
	varFcbaName = ""
	varCompanyName = ""
	varCompanyCode = ""
End If
rsBaInfo.Close
set rsBaInfo = Nothing

qSQL = " SELECT fccode, fcname, parent_fccode, priority, status" & _
       " FROM setup_bisa_agrotech_tab" & _
       " ORDER BY SUBSTR(fccode, 1, 2), fccode"
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open qSQL, connstr, adOpenStatic

arr_setup_eof = true
If Not qORS.EOF Then
	arr_setup = qORS.getrows()
	arr_setup_eof = false
end if
qORS.Close
set qORS = Nothing

Function f_colspan(data_array__, code__)
	result_ = 0
	arr_1 = data_array__
	for indx = 0 to UBound(arr_1, 2)
		if len("" & code__) = 2 and len("" & arr_1(2, indx)) = 2 then			
			if cstr(arr_1(2, indx)) = cstr(code__) then
				result_ = result_ +1
			end if
		end if
	next
	f_colspan = result_
End Function

Function f_maxrow_indeks(data_array__, indeks__)
	result_ = 0
	arr_1 = data_array__
	for indx = 0 to UBound(arr_1, 2)
		if cstr("" & arr_1(2, indx)) = cstr("" & indeks__) then
			result_ = result_ +1
		end if
	next
	f_maxrow_indeks = result_
End Function

Function f_totalinput_indeks(data_array__, indeks__)
	result_ = 0
	arr_1 = data_array__
	for indx = 0 to UBound(arr_1, 2)
		if Left(cstr("" & arr_1(0, indx)),2) = cstr("" & indeks__) then
			if cdbl(to_nol(arr_1(2, indx))) = 1 then
				result_ = result_ +1
			end if
		end if
	next
	f_totalinput_indeks = result_
End Function


if not arr_setup_eof then
	for yy=0 to UBound(arr_setup,2)
		if len(cstr("" & arr_setup(0,yy))) = 4 and cstr("" & arr_setup(4,yy)) = txtstatus then
			maxCol = maxCol +1
			d_span = d_span +1
		end if		
	next
end if


%>

<%
			Dim strTitle, rsTitle
			Dim titleGroupPlantation, titlePlantation, titleGroupRegion, titleRegion, titleCompany, titleSite
			
	        strTitle = " select PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPPLANTATION') titleGP, " & _
							    " PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','PLANTATION') titleP , " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPREGION') titleGR, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','REGION') titleR, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','COMPANY') titleC, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','SITE') titleS from dual"
								
			Set rsTitle = Server.CreateObject("ADODB.Recordset")
			'response.write rsTitle
			rsTitle.Open strTitle, connstr, 3,3	

			if not rsTitle.EOF then
				no = 0
				titleGroupPlantation = rsTitle ("titleGP")			
				titlePlantation = rsTitle ("titleP")
				titleGroupRegion = rsTitle ("titleGR")
				titleRegion = rsTitle ("titleR")
				titleCompany = rsTitle ("titleC")
				titleSite = rsTitle ("titleS")
			else
				titleGroupPlantation = "Group Plantation"			
				titlePlantation = "Plantation"
				titleGroupRegion = "Group Region"
				titleRegion = "Region"
				titleCompany = "Company"
				titleSite = "Site"
			end if
			rsTitle.Close
			set rsTitle = Nothing
			
%>
<html>

<head>
<title>Fire Danger Rating - iPlaS Reporting</title>
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
	$( ".datepicker" ).datepicker({
		changeMonth: true,
		changeYear: true
	});
});
</script>
<!--
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
-->

<style type="text/css">
@media print
{
.printhide{display:none;}
}

th.rotate-270 {
  height: 200px; /* 1 */
  white-space: nowrap;
  min-width:40px;
  max-width:40px;
  overflow:hidden;
}
th.rotate-270 > div {
	transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
	-webkit-transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
	-ms-transform: translate(8px /* 2 */ , 80px /* 3 */ ) rotate(-90deg);
    width:15px; /* 4 */
	font-size:10px;
	font-weight:normal;
}

/*
1. Mengatur tinggi table head.
2. mengatur posisi X div
3. mengatur posisi Y div.
4. mengatur lebarnya DIV yang akan mempengaruhi Lebarnya Table Head.
*/
</style>


<link href="/plugins/chosen/chosen.min.css" rel="stylesheet" />
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
						<td><font face="verdana" size="1">TAHUN</font></td>
					<% 
						Dim mon, Yr, MinYr, MaxYr, i, MM

						Yr = txtYear
						MinYr = 2015
						MaxYr = Yr  + 2
					%>
							<td width="4">
								<select name="txtYear">
					<%
						For i = MinYr to MaxYr
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(txtYear) then%> Selected <%End If%>><%=i%></option>
					<%
						Next
					%>
								</select>
							</td>
							<td><font face="verdana" size="1">BULAN</font></td>
							<td width="2">
								<select name="txtMonth">
					<%
						For i = 1 to 12
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
					<%
						Next
					%>
								</select>
							</td>
							<td>
								<font face="verdana" size="1"><%=titleSite%></font>
							</td>
							<td>
								<select name="txtFcba" >
									<%
									Dim rsFcba, strFcba
									strFcba	=	" SELECT FCCODE, ORDERDATA FROM BUSINESSUNIT" & _
												" WHERE USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL" & _
												" AND FCTYPE = 'E'" & _
												" ORDER BY ORDERDATA"
									
									Set rsFcba = Server.CreateObject("ADODB.Recordset")
									rsFcba.Open strFcba, connstr, adOpenStatic
									
									If rsFcba.Bof Or rsFcba.eof Then
									%>
									<option value=""></option>
									<%
									End If
									
									if not rsFcba.eof then
									Do While Not rsFcba.eof
									%>
									<option value="<%=rsFcba("FCCODE")%>" <%if v_txtFcba = rsFcba("FCCODE") then%> Selected <%End If%>><%=rsFcba("FCCODE")%></option>
									<%
										rsFcba.movenext
									loop
									end if
									rsFcba.Close
									set rsFcba = Nothing
									%>
								</select>
							</td>
							<td>
								<font face="verdana" size="1">AFDELING</font>
							</td>
							<td>
								<select name="txtAfdeling" >
									<option value="">--All--</option>
									<%
									Dim rsAfdeling, strAfdeling
									strAfdeling	= " SELECT fccode, fcname FROM parameter WHERE fcba = '" & v_txtFcba & "' AND parcode = 'F0001'" & _
												  " UNION ALL" & _
												  " SELECT fccode, fcname FROM parameter WHERE fcba = '" & v_txtFcba & "' AND parcode = 'OMBROMETER_LOCATION'"
									
									Set rsAfdeling = Server.CreateObject("ADODB.Recordset")
									rsAfdeling.Open strAfdeling, connstr, adOpenStatic 
									
									if not rsAfdeling.eof then
									Do While Not rsAfdeling.eof
									if rsAfdeling("FCCODE") = v_txtAfdeling then%>
									<option value="<%=rsAfdeling("FCCODE")%>" Selected><%=rsAfdeling("FCCODE")%></option> 
									<%Else%>
									<option value="<%=rsAfdeling("FCCODE")%>"><%=rsAfdeling("FCCODE")%></option>  
									<%End if
										rsAfdeling.movenext
									loop
									end if
									rsAfdeling.Close
									set rsAfdeling = Nothing
									%>
								</select>
							</td>
							<td><input type="submit" value="View" /></td>
						</tr>				
					</table>
				</div>
			</form>
		</div>
<table border="0" width="100%" id="product-table" class="xthf-site" >
	<thead style="border-left:1px #000000 solid;">
	<tr>
         <th colspan="<%=d_span%>" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
		 <div id="header-wrapper">
			<div id="head_judul">
				<h4><%If Len(Trim(varCompanyName)) > 0 Then%>PT. <%End If%><%=varCompanyName%></h4>
				<h5><%=varFcbaName%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>Fire Danger Rating</h2>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(txtMonth)%>&nbsp;<%=txtYear%></h4>
			</div>
          </div>  
		 </th>
	</tr>
	<tr>
		<th>Ombrometer</th>
		<th>Date</th>
		<th>1400 Relative Humidity (%)</th>
		<th>RH</th>
		<th>Daily Rainfall (mm)</th>
		<th>Rainless Days</th>
		<th>Total Rainfall in Last 15 Days (mm)</th>
		<th>Fuel Conditions</th>
		<th>RD</th>
		<th>TRL15</th>
		<th>FC</th>
		<th>FDI</th>
		<th>FDR</th>
		<th>PENUNJUKKAN WARNA ANAK PANAH di PAPAN MONITORING</th>
	</tr>
    </thead>
	<tbody>
<%
if not arr_field_eof then
	runIndex = 0
	maxIndeks = 0
	inputIndeks = 0
	currIndeks = ""
	prevIndeks = ""
	
	st_jan = 0
	st_feb = 0
	st_mar = 0
	st_apr = 0
	st_may = 0
	st_jun = 0
	st_jul = 0
	st_aug = 0
	st_sep = 0
	st_oct = 0
	st_nov = 0
	st_dec = 0

	for y=0 to ubound(arr_field,2)
		runIndex = runIndex +1
		currIndeks = arr_field(2,y)
		
		response.write "<tr>"
		if currIndeks <> prevIndeks then
			maxIndeks = f_maxrow_indeks(arr_field, currIndeks)
			response.write "	<td rowspan='"&maxIndeks&"' valign='top' align='center'>"&arr_field(2,y)&"</td>"
		end if
		response.write "	<td align='center'>"&Day(arr_field(1,y))&"</td>"
		response.write "	<td align='center'>"&arr_field(3,y)&"</td>"
		response.write "	<td align='right'>"&arr_field(4,y)&"</td>"
		response.write "	<td align='right'>"&arr_field(5,y)&"</td>"
		response.write "	<td align='right'>"&arr_field(6,y)&"</td>"
		response.write "	<td align='right'>"&arr_field(7,y)&"</td>"
		response.write "	<td>"&arr_field(8,y)&"</td>"
		response.write "	<td align='right'>"&arr_field(9,y)&"</td>"
		response.write "	<td align='right'>"&arr_field(10,y)&"</td>"
		response.write "	<td align='right'>"&arr_field(11,y)&"</td>"
		response.write "	<td align='right'>"&arr_field(12,y)&"</td>"
		response.write "	<td>"&arr_field(13,y)&"</td>"
		response.write "	<td>"&arr_field(14,y)&"</td>"
		
			
		
		if runIndex = maxIndeks then			
			runIndex = 0
		end if
		
		response.write "</tr>"
		prevIndeks = currIndeks
	next
else
	response.write "<tr><td colspan='"&d_span&"'>NO DATA AVAILABLE</td></tr>"
end if
%>
	</tbody>
</table>


<script src="/plugins/chosen/chosen.jquery.min.js"></script>
	<script type="text/javascript">	
		var config = {
		  '.chosen-select'           : {},
		  '.chosen-select-deselect'  : {allow_single_deselect:true},
		  '.chosen-select-no-single' : {disable_search_threshold:10},
		  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
		  '.chosen-select-width'     : {width:"95%"}
		}
		for (var selector in config) {
		  $(selector).chosen(config[selector]);
		}
		
	</script>
	</body>
</html>