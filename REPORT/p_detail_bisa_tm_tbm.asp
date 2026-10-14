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
txtstatus = Request("txtstatus")

If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))
If "" & txtMonth = "" Then txtMonth = Month(Session("v_Costbook"))
If "" & txtstatus = "" Then txtstatus = "TM"

add_col_bud_prod_qty = ""
if txtstatus = "TM" then
	d_span = 11
	add_col_bud_prod_qty = ", hdr.bgt_prod_qty, hdr.act_prod_qty"
else
	d_span = 8
	add_col_bud_prod_qty = ""
end if
	

v_txtAfdeling	=	Request("txtAfdeling")
v_txtBlok		=	Request("txtBlok")
v_txtFcba		=	Request("txtFcba")

Dim v_txtPersen : v_txtPersen = Request("txtPersen")
If Len(trim(v_txtPersen)) = 0 Then v_txtPersen = 0


strDefaultBa = " SELECT" & _
				"     FCCODE as FCBA" & _
				" FROM BUSINESSUNIT BU" & _
				" INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) " & _
				" WHERE" & _
				"     ISDEFAULT = 'TRUE'" & _
				"     AND FCTYPE = 'E'" & _
				"     AND FCUSERCODE = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("FCBA")
Else
	varDefaultBa = ""
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
	v_txtFcba = varDefaultBa
End If

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")


arr_data_eof = true
strPusingan	=	" SELECT *" & _
				" FROM trans_bisa_agrotech_dtl" & _
				" WHERE fcba = '" & v_txtFcba & "'" & _
				"       AND accountingyear = '" & txtYear & "'" & _
				"       AND accountingperiod = '" & txtMonth & "'"
'response.write strPusingan
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open strPusingan, connstr, adOpenStatic

If Not qORS.EOF Then
	arr_data = qORS.getrows()
	arr_data_eof = false
End If
qORS.Close
set qORS = Nothing

Function f_value(data_array__, fieldcode__, fccode__)
	result_ = 0
	arr_1 = data_array__
	if not arr_data_eof then
	for indx = 0 to UBound(arr_1, 2)
		if "" & fieldcode__ <> "" and "" & fccode__ <> "" then			
			if cstr(arr_1(2, indx)) = cstr(fieldcode__) and cstr(arr_1(3, indx)) = cstr(fccode__) then
				result_ = cdbl(to_nol(arr_1(4, indx)))
			end if
		end if
	next
	end if
	f_value = result_
End Function

fddate = txtMonth&"/1/"&txtYear

if "" & v_txtBlok <> "" then
	wblok = v_txtBlok
end if

arr_field_eof = true
strPusingan	= " SELECT hdr.fieldcode" & _
              "        , hdr.plantingyear" & _
              "        , hdr.hectarageplanted" & _
              "        , hdr.totalstand" & _
              "        , hdr.plantingmaterial" & _
              "        , hdr.soiltype" & _
              "        , hdr.terraintype" & _
              "        , hdr.bgt_prod_qty" & _
              "        , hdr.act_prod_qty" & _
              "        , hdr.total_prioritas1" & _
              "        , hdr.total_prioritas2" & _
              "        , hdr.total_prioritas3" & _
              "        , hdr.division" & _
              "        , CASE WHEN hdr.hectarageplanted = 0 THEN 0 ELSE ROUND(hdr.totalstand/hdr.hectarageplanted,0) END sph" & _
              "        , hdr.nilai_akhir" & _
              "        , hdr.skor_akhir" & _
              "        "&add_col_bud_prod_qty&"" & _
              " FROM trans_bisa_agrotech_hdr hdr" & _
			  " LEFT JOIN (" & _
			  "     SELECT DISTINCT dtl.fcba, dtl.accountingyear, dtl.accountingperiod, dtl.fieldcode" & _
			  "     FROM trans_bisa_agrotech_dtl dtl" & _
			  "     INNER JOIN setup_bisa_agrotech_tab stp ON dtl.fccode = stp.fccode" & _
			  "     WHERE dtl.fcba = '" & v_txtFcba & "'" & _
			  "           AND dtl.accountingyear = '" & txtYear & "'" & _
			  "           AND dtl.accountingperiod = '" & txtMonth & "'"
if "" & txtstatus <> "" then
	strPusingan = strPusingan & "           AND stp.status = '" & txtstatus & "'"
end if

if "" & v_txtBlok <> "" then
	strPusingan = strPusingan & "           AND dtl.fieldcode = '" & v_txtBlok & "'"
end if
strPusingan = strPusingan & " ) dtl ON ( dtl.fcba = hdr.fcba" & _
                            "            AND dtl.accountingyear = hdr.accountingyear" & _
                            "            AND dtl.accountingperiod = hdr.accountingperiod" & _
                            "            AND dtl.fieldcode = hdr.fieldcode)" & _
                            " WHERE hdr.fcba = '" & v_txtFcba & "'" & _
                            "       AND hdr.accountingyear = '" & txtYear & "'" & _
                            "       AND hdr.accountingperiod = '" & txtMonth & "'" & _
                            "       AND hdr.status = '" & txtstatus & "'"

if "" & v_txtAfdeling <> "" then
	strPusingan = strPusingan & "       AND hdr.division = '" & v_txtAfdeling & "'"
end if


strPusingan = strPusingan & " ORDER BY division, fieldcode"
'response.write strPusingan
Set qORS = Server.CreateObject("ADODB.Recordset")
qORS.Open strPusingan, connstr, adOpenStatic

If Not qORS.EOF Then
	arr_field = qORS.getrows()
	arr_field_eof = false
End If
qORS.Close
set qORS = Nothing

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

Function f_maxrow_division(data_array__, code__)
	result_ = 0
	arr_1 = data_array__
	for indx = 0 to UBound(arr_1, 2)
		if cstr("" & arr_1(12, indx)) = cstr("" & code__) then
			result_ = result_ +1
		end if
	next
	f_maxrow_division = result_
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
<title>BISA Agrotechnical - iPlaS Reporting</title>
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
								<font face="verdana" size="1">STATUS</font>
							</td>
							<td>
								<select name="txtstatus" >
									<option value="TM" <%if txtstatus = "TM" then%> selected <%end if%>>TM</option>
									<option value="TBM" <%if txtstatus = "TBM" then%> selected <%end if%>>TBM</option>
								</select>
							<td>
								<font face="verdana" size="1">AFDELING</font>
							</td>
							<td>
								<select name="txtAfdeling" >
									<option value="">--ALL--</option>
									<%
									Dim rsAfdeling, strAfdeling
									strAfdeling	=	" SELECT FCCODE FROM PARAMETER WHERE PARCODE = 'F0001' AND FCBA = '" & v_txtFcba & "' ORDER BY FCCODE"
									
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
							<td>
								<font face="verdana" size="1">BLOK</font>
							</td>
							<td>
								<select name="txtBlok" >
									<option value="">--ALL--</option>
									<%
									Dim rsBlok, strBlok
									strBlok	=	" SELECT fccode, fcname FROM field WHERE fcba = '" & v_txtFcba & "'" & _
												" AND field_api.getvalid_isactive_harvested(fccode, fcba, TO_DATE('"&v_txtDate&"', 'mm/dd/yyyy')) = 'TRUE' ORDER BY fccode"
									
									Set rsBlok = Server.CreateObject("ADODB.Recordset")
									rsBlok.Open strBlok, connstr, adOpenStatic 
									
									if not rsBlok.eof then
									Do While Not rsBlok.eof
									if rsBlok("FCCODE") = v_txtBlok then%>
									<option value="<%=rsBlok("FCCODE")%>" Selected><%=rsBlok("FCCODE")%> - <%=rsBlok("FCNAME")%></option> 
									<%Else%>
									<option value="<%=rsBlok("FCCODE")%>"><%=rsBlok("FCCODE")%> - <%=rsBlok("FCNAME")%></option>  
									<%End if
										rsBlok.movenext
									loop
									end if
									rsBlok.Close
									set rsBlok = Nothing
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
				<h2>BEST IMPROVEMENT SELF ASSESSMENT (BISA) <%=txtstatus%></h2>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(txtMonth)%>&nbsp;<%=txtYear%></h4>
			</div>
          </div>  
		 </th>
	</tr>
	<tr>
		<th rowspan="2">Blok</th>
		<th rowspan="2">TT</th>
		<th rowspan="2">Luas (Ha)</th>
		<th rowspan="2">SPH (PKK/Ha)</th>
		<th rowspan="2">Jenis Bibit</th>
		<th rowspan="2">Jenis Tanah</th>
		<th rowspan="2">Topografi</th>
		<%if txtstatus = "TM" then%>
		<th rowspan="2">Bgt Produksi (Ton/Ha)</th>
		<th rowspan="2">Real Produksi (Ton/Ha)</th>
		<th rowspan="2">Ach Produksi (%)</th>
		<%end if%>
		<%if not arr_setup_eof then%>
			<%for nn=0 to ubound(arr_setup,2)%>
				<%if len(cstr("" & arr_setup(0,nn))) = 2 and cstr("" & arr_setup(4,nn)) = txtstatus then%>
					<th colspan="<%=f_colspan(arr_setup, arr_setup(0,nn))%>"><%=arr_setup(1,nn)%></th>
				<%end if%>
			<%next%>
		<%end if%>
		<th rowspan="2" style="min-width:50px;max-width:50px;">Nilai Akhir Kualitas Blok</th>
	</tr>
	<tr>
		<%if not arr_setup_eof then%>
			<%for nn=0 to ubound(arr_setup,2)%>
				<%if len(cstr("" & arr_setup(0,nn))) = 4 and cstr("" & arr_setup(4,nn)) = txtstatus then%>
				<%
					if cdbl(to_nol(arr_setup(3,nn))) = 1 then
						th_style = "background-color:#FF4136;color:#FFFFFF;"
					elseif cdbl(to_nol(arr_setup(3,nn))) = 2 then
						th_style = "background-color:#FFDC00;color:#111111;"
					elseif cdbl(to_nol(arr_setup(3,nn))) = 3 then
						th_style = "background-color:#2ECC40;color:#111111;"
					end if
				%>
					<th class="rotate-270" style="<%=th_style%>"><div><%=arr_setup(1,nn)%></div></th>
				<%end if%>
			<%next%>
		<%end if%>
	</tr>

    </thead>
	<tbody>
<%
	if not arr_field_eof then
		defST = 0
		st_ha = 0
		st_pokok = 0
		
		st_biru = 0
		st_kuning = 0
		st_hijau = 0
		st_merah = 0
		
		st_skor = 0
		st_skor_pembagi = 0
		
		for z=1 to maxCol
			execute("bc" & z & " = 0")
			execute("bnc" & z & " = 0")
		next
		
		for y=0 to ubound(arr_field,2)
			bud_ton_ha = 0
			act_ton_ha = 0
			ach = 0
		
			defST = defST +1
			cDivision = arr_field(12,y)
			if cDivision <> pDivision then
				maxST = f_maxrow_division(arr_field, cDivision)
			end if
			
			st_ha = st_ha + cdbl(to_nol(arr_field(2,y)))
			st_pokok = st_pokok + cdbl(to_nol(arr_field(3,y)))
			
			response.write "<tr>"
			response.write "	<td align='left'>"&arr_field(0,y)&"</td>"
			response.write "	<td align='center'>"&arr_field(1,y)&"</td>"
			response.write "	<td align='right'>"&arr_field(2,y)&"</td>"
			response.write "	<td align='right'>"&arr_field(13,y)&"</td>"
			response.write "	<td align='left'>"&arr_field(4,y)&"</td>"
			response.write "	<td align='left'>"&arr_field(5,y)&"</td>"
			response.write "	<td align='left'>"&arr_field(6,y)&"</td>"
			
			if txtstatus = "TM" then
				if cdbl(to_nol(arr_field(2,y))) = 0 then
					bud_ton_ha = 0
					act_ton_ha = 0
				else
					bud_ton_ha = Round(cdbl(to_nol(arr_field(16,y)))/cdbl(to_nol(arr_field(2,y))),1)
					act_ton_ha = Round(cdbl(to_nol(arr_field(17,y)))/cdbl(to_nol(arr_field(2,y))),1)
				end if
				
				if bud_ton_ha = 0 then
					if act_ton_ha > 0 then
						ach = 100
					else
						ach = 0
					end if
				else
					ach = Round(act_ton_ha/bud_ton_ha*100, 0)
				end if
				
				if ach >= 100 then
					style_prod = "background-color:#0074D9;color:#FFFFFF;"
				elseif ach >= 60 and ach < 100 then
					style_prod = "background-color:#FFDC00;"
				elseif ach < 60 then
					style_prod = "background-color:#FF4136;"
				else
					style_prod = "background-color:#FFFFFF;"
				end if
				
				response.write "	<td align='right'>"&bud_ton_ha&"</td>"
				response.write "	<td align='right'>"&act_ton_ha&"</td>"
				response.write "	<td align='right' style='"&style_prod&"'>"&ach&"</td>"
			end if
			if not arr_setup_eof then
				for yy=0 to ubound(arr_setup,2)
					if len(cstr("" & arr_setup(0,yy))) = 4 and cstr("" & arr_setup(4,yy)) = txtstatus then
						s_style = ""
						if f_value(arr_data, arr_field(0,y), arr_setup(0,yy) ) = 1 then
							p_value = "x"
							s_style = s_style & "color:#FF4136;font-weight:bold;"
							
						else
							p_value = "v"
							s_style = s_style & "color:#0074D9;font-weight:bold;"
							
						end if
						response.write "	<td align='center' style='"&s_style&"'>"&p_value&"</td>"
					end if
				next
			end if
			
			if not arr_setup_eof then
				for z=1 to maxCol
					yyz = 0
					for yy=0 to ubound(arr_setup,2)
						if len(cstr("" & arr_setup(0,yy))) = 4 and cstr("" & arr_setup(4,yy)) = txtstatus then
							yyz = yyz +1
							if f_value(arr_data, arr_field(0,y), arr_setup(0,yy) ) = 1 then
								if z = yyz then execute("bnc" & z & " = bnc" & z & " +1") end if
							else								
								if z = yyz then execute("bc" & z & " = bc" & z & " +1") end if
							end if
						end if
					next
				next
			end if
			
			nilai_akhir = ""
			skor_akhir = 0
			
			nilai_akhir = cstr(""&arr_field(14,y))
			skor_akhir = cdbl(to_nol(arr_field(15,y)))
			
			if skor_akhir > 0 then
				st_skor = st_skor +skor_akhir
				st_skor_pembagi = st_skor_pembagi +1
			end if
			
			if nilai_akhir = "A" then
				n_style = "background-color:#0074D9;color:#0074D9;"
				st_biru = st_biru +1
			elseif nilai_akhir = "B" then
				n_style = "background-color:#2ECC40;color:#2ECC40;"
				st_hijau = st_hijau +1
			elseif nilai_akhir = "C" then
				n_style = "background-color:#FFDC00;color:#FFDC00;"
				st_kuning = st_kuning +1
			elseif nilai_akhir = "D" then
				n_style = "background-color:#FFDC00;color:#FFDC00;"
				st_kuning = st_kuning +1
			elseif nilai_akhir = "E" then
				n_style = "background-color:#FF4136;color:#FF4136;"
				st_merah = st_merah +1
			else
				n_style = "background-color:white;color:black;"
			end if
			
			response.write "	<td align='center' style='"&n_style&"'>"&nilai_akhir&"</td>"
			response.write "</tr>"
			
			if defST = maxST then
				if txtstatus = "TM" then
					span_td = 6
				else
					span_td = 3
				end if
				
				response.write "<tr>"
				response.write "	<td align='center' rowspan='2' colspan='2'> Total "&cDivision&"</td>"
				response.write "	<td align='right' rowspan='2' style='font-weight:bold;'>"&st_ha&"</td>"
				if st_ha = 0 then
					response.write "	<td align='right' rowspan='2'>0</td>"
				else
					response.write "	<td align='right' rowspan='2' style='font-weight:bold;'>"&RoundUp(st_pokok/st_ha,0)&"</td>"
				end if
				response.write "	<td align='center' colspan='"&span_td&"'> Blok Conformance</td>"
				for z=1 to maxCol
					response.write "	<td align='right' style='font-weight:bold;'> "&eval("bc"&z)&"</td>"
				next
				response.write "	<td align='center'>&nbsp;</td>"
				response.write "</tr>"
				
				response.write "<tr>"
				response.write "	<td align='center' colspan='"&span_td&"'> Blok Non Conformance</td>"
				for z=1 to maxCol
					response.write "	<td align='right' style='font-weight:bold;'> "&eval("bnc"&z)&"</td>"
				next
				response.write "	<td align='center'>&nbsp;</td>"
				response.write "</tr>"
				
				response.write "<tr>"
				response.write "	<td align='center' colspan='"&d_span -1&"' style='border-bottom:none;border-right:none;'></td>"
				response.write "	<td align='center' style='border-left:none;border-bottom:none;'></td>"
				response.write "</tr>"
				response.write "<tr>"
				response.write "	<td align='center' colspan='"&d_span -8&"' style='border-top:none;border-bottom:none;border-right:none;'></td>"
				response.write "	<td align='center' style='background-color:#0074D9;border:none;'>&nbsp;</td>"
				response.write "	<td align='center' style='border:none;font-weight:bold;'>"&st_biru&"</td>"
				response.write "	<td align='left' style='border:none;font-weight:bold;'>Blok</td>"
				response.write "	<td align='left' style='border:none;'>&nbsp;</td>"
				response.write "	<td align='center' style='background-color:#FFDC00;border:none;'>&nbsp;</td>"
				response.write "	<td align='center' style='border:none;font-weight:bold;'>"&st_kuning&"</td>"
				response.write "	<td align='left' style='border:none;font-weight:bold;'>Blok</td>"
				response.write "	<td align='center' style='border-top: 1px solid #3A5F0B;font-weight:bold;'>Score</td>"
				response.write "</tr>"
				
				if st_skor_pembagi = 0 then
					st_skor_akhir = 0
				else
					st_skor_akhir = RoundUp(st_skor/st_skor_pembagi,0)
				end if
				
				if st_skor_akhir < 75 then
					st_n_style = "font-weight:bold;background-color:#FF4136;color:#FFFFFF;"
				elseif st_skor_akhir >= 75 and st_skor_akhir < 85 then
					st_n_style = "font-weight:bold;background-color:#FFDC00;color:#111111;"					
				elseif st_skor_akhir >= 85 and st_skor_akhir < 100 then
					st_n_style = "font-weight:bold;background-color:#2ECC40;color:#111111;"
				else
					st_n_style = "font-weight:bold;background-color:#0074D9;color:#FFFFFF;"
				end if
				
				response.write "<tr>"
				response.write "	<td align='center' colspan='"&d_span -1&"' style='border-top:none;border-bottom:none;'></td>"
				response.write "	<td align='center' rowspan='2' style='"&st_n_style&"'>"&st_skor_akhir&"</td>"
				response.write "</tr>"				
				response.write "<tr>"
				response.write "	<td align='center' colspan='"&d_span -8&"' style='border-top:none;border-bottom:none;border-right:none;'></td>"
				response.write "	<td align='center' style='background-color:#2ECC40;border:none;'>&nbsp;</td>"
				response.write "	<td align='center' style='border:none;font-weight:bold;'>"&st_hijau&"</td>"
				response.write "	<td align='left' style='border:none;font-weight:bold;'>Blok</td>"
				response.write "	<td align='left' style='border:none;'>&nbsp;</td>"
				response.write "	<td align='center' style='background-color:#FF4136;border:none;'>&nbsp;</td>"
				response.write "	<td align='center' style='border:none;font-weight:bold;'>"&st_merah&"</td>"
				response.write "	<td align='left' style='border:none;font-weight:bold;'>Blok</td>"
				response.write "</tr>"
				
				response.write "<tr><td align='center' colspan='"&d_span&"' style='border-top:none;'></td></tr>"
				
				if y <> ubound(arr_field,2) then
					response.write "<tr><td align='center' colspan='"&d_span&"'></td></tr>"
				end if
				defST = 0
				for z=1 to maxCol
					execute("bc" & z & " = 0")
					execute("bnc" & z & " = 0")
				next
				st_ha = 0
				st_pokok = 0
				
				st_biru = 0
				st_kuning = 0
				st_hijau = 0
				st_merah = 0
				
				st_skor = 0
				st_skor_pembagi = 0
			end if
			
			pDivision = cDivision
		next
else
	response.write "<tr><td colspan='"&d_span&"'>NO DATA AVAILABLE</td></tr>"
end if
%>
</tbody>
</table>




<style>
.stick { display:none;}
.xfTBarF { height:15px; line-height: 15px; font-size:10px; padding-left:20px; background-color:#008001;}
.xfContainer { border:1px #008001 solid; box-shadow:1px 1px 1px rgba(0,0,0,0.5); background:#FFFFFF;}
.fenster-content { padding:5px; padding-top:10px;}
#indicator_color { border:1px #999999 solid ; width:650px; padding:2px; font-family:Arial, Helvetica, sans-serif; }
#indicator_color table { border-collapse:collapse; border-spacing:0px;}
#indicator_color table, #indicator_color tr, #indicator_color td { border: none; line-height:30px; }
#indicator_color th { font-size:12px; background-color:#222222; border-left:1px #666666 dashed; color:#FFFFFF;}
#indicator_color .2nd { background-color:#f1f1f1;}
#indicator_color td { font-size:12px; }
.merah { width:15px; height:15px; background-color:#FF0000; float:left; margin-right:5px; margin-top:7px;
   	      border-radius:3px;
		  -moz-border-radius:3px;
		  -webkit-border-radius:3px;
		  -o-border-radius:3px;
		  box-shadow:1px 1px 0px #ccc;
		  -moz-box-shadow:1px 1px 0px #ccc;
		  -webkit-box-shadow:1px 1px 0px #ccc;
		  -o-box-shadow:1px 1px 0px #ccc;}
.kuning { width:15px; height:15px; background-color:#FFFF00; float:left; margin-right:5px; margin-top:7px;
		   border-radius:3px;
		  -moz-border-radius:3px;
		  -webkit-border-radius:3px;
		  -o-border-radius:3px;
		  box-shadow:1px 1px 0px #ccc;
		  -moz-box-shadow:1px 1px 0px #ccc;
		  -webkit-box-shadow:1px 1px 0px #ccc;
		  -o-box-shadow:1px 1px 0px #ccc;}
.hijau { width:15px; height:15px; background-color:#008001; float:left; margin-right:5px; margin-top:7px;
          border-radius:3px;
		  -moz-border-radius:3px;
		  -webkit-border-radius:3px;
		  -o-border-radius:3px;
		  box-shadow:1px 1px 0px #ccc;
		  -moz-box-shadow:1px 1px 0px #ccc;
		  -webkit-box-shadow:1px 1px 0px #ccc;
		  -o-box-shadow:1px 1px 0px #ccc;}
#indicator_color .span_color { float:left;}
#indicator_color .point_table {}
#indicator_color .MsoNormal { margin-bottom:3px; border-bottom:1px #ccc dashed; padding:0px 5px;}
.symbol_1 { font-size:18px; color:#008001; margin-right:5px; }
</style>

<div id="indicator_color" style="margin-bottom:5px;">
	<div style="float:left;margin-right:20px;">
	<table cellpadding="0" border="1" style="border:none">
		<thead>
			<tr style="height:40px;">
				<th style=" border-left:none;" width="319" valign="center" colspan="2">
					<p>Kriteria Data</p>
				</th>
			</tr>
		</thead>
		<tbody>
			<tr bgcolor="#f1f1f1">
				<td align="left" valign="center" style="border-top:none;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;width:35%;">
					<p class="MsoNormal">Prioritas I</p>
				</td>
				<td align="left" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;">
					<p style="background-color:#FF4136;color:#FFFFFF;padding-left:10px;">Sample Text</p>
				</td>
				
				
			</tr>
			
			<tr>
				<td align="left" valign="center" style="border-top:none;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">Prioritas I I</p>
				</td>
				<td align="left" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;">
					<p style="background-color:#FFDC00;padding-left:10px;">Sample Text</p>
				</td>
				
			</tr>
			
			<tr bgcolor="#f1f1f1">
				<td align="left" valign="center" style="border-top:none;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">Prioritas I I I</p>
				</td>
				<td align="left" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;">
					<p style="background-color:#2ECC40;padding-left:10px;">Sample Text</p>
				</td>
				
			</tr>
		</tbody>
	</table>
	</div>

	<div>
	<table cellpadding="0" border="1" style="border:none">
		<thead>
			<tr>
				<th style=" border-left:none;" width="319" valign="center" colspan="6">
					<p>Kriteria Nilai<br/>Jumlah Non Conformance Minimal</p>
				</th>
			</tr>
			<tr>
				<th valign="center" style="min-width:100px;max-width:100px;"><p>Prioritas</p></th>
				<th valign="center" style="min-width:40px;max-width:40px;"><p>A</p></th>
				<th valign="center" style="min-width:40px;max-width:40px;"><p>B</p></th>
				<th valign="center" style="min-width:40px;max-width:40px;"><p>C</p></th>
				<th valign="center" style="min-width:40px;max-width:40px;"><p>D</p></th>
				<th valign="center" style="min-width:40px;max-width:40px;"><p>E</p></th>
			</tr>
		</thead>
		<tbody>
			<tr bgcolor="#f1f1f1">
				<td align="center" valign="center" style="border-top:none;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">I</p>
				</td>
				<td align="center" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">0</p>
				</td>
				
				<td align="center" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">0</p>
				</td>
				
				<td align="center" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">0</p>
				</td>
				
				<td align="center" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">0</p>
				</td>
				
				<td align="center" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">1</p>
				</td>
			</tr>
			
			<tr>
				<td align="center" valign="center" style="border-top:none;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">I I</p>
				</td>
				<td align="center" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">0</p>
				</td>
				
				<td align="center" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">0</p>
				</td>
				
				<td align="center" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">1</p>
				</td>
				
				<td align="center" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">1</p>
				</td>
				
				<td align="center" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">0</p>
				</td>
			</tr>
			
			<tr bgcolor="#f1f1f1">
				<td align="center" valign="center" style="border-top:none;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">I I I</p>
				</td>
				<td align="center" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">0</p>
				</td>
				
				<td align="center" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">1</p>
				</td>
				
				<td align="center" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">0</p>
				</td>
				
				<td align="center" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">1</p>
				</td>
				
				<td align="center" class="point_table" valign="top" style="border-top:none;border-left:none;border-right:solid white 1.0pt;padding:0in 5.4pt 0in 5.4pt; font-weight:bold;">
					<p class="MsoNormal">0</p>
				</td>
			</tr>
		</tbody>
	</table>
	</div>
</div>

	</body>
</html>