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

txtcompany = Request("txtcompany")
txtreport = Request("txtreport")
v_txtFcba = Request("txtFcba")
v_txtYear = Request("txtYear")
v_txtMonth = Request("txtMonth")

If "" & v_txtYear = "" Then v_txtYear = Year(Session("v_Costbook"))
If "" & v_txtMonth = "" Then v_txtMonth = Month(Session("v_Costbook"))

strDefaultBa = " SELECT fccode AS fcba, fccompanycode" & _
               " FROM BUSINESSUNIT BU" & _
               " INNER JOIN user_fcba_tab uft ON (bu.fccode = uft.fcba) " & _
               " WHERE isdefault = 'TRUE'" & _
               "       AND fcusercode = '" & Session("CodeName") & "'"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	varDefaultBa = rsDefaultBa("FCBA")
	varcompanycode = rsDefaultBa("fccompanycode")
Else
	varDefaultBa = ""
	varcompanycode = ""
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

If "" & v_txtFcba = "" Then v_txtFcba = varDefaultBa
If "" & txtcompany = "" Then txtcompany = varcompanycode
If "" & txtreport = "" Then txtreport = "HUTANGSUPPLIER"

ScriptName = Request.ServerVariables("SCRIPT_NAME")

arr_param_eof = true
strDefaultBa = " SELECT fccode, fcname" & _
               " FROM parameter" & _
               " WHERE fcba = 'GWE'" & _
               "       AND parcode = 'RPT_HUTANGPIUTANG'" & _
			   " ORDER BY rate_1"

Set rsDefaultBa = Server.CreateObject("ADODB.Recordset")
rsDefaultBa.Open strDefaultBa, connstr, adOpenStatic

If Not rsDefaultBa.EOF Then
	arr_param = rsDefaultBa.getrows()
	arr_param_eof = false
End If
rsDefaultBa.Close
set rsDefaultBa = Nothing

if txtreport = "HUTANGSUPPLIER" then
	tablename = "apsupplier"
	jointable1 = "supplier"
elseif txtreport = "PINJAMANKARYAWAN" then
	tablename = "employeeloan"
	jointable1 = "employee"
elseif txtreport = "PINJAMANKONTRAKTOR" then
	tablename = "contractorloan"
	jointable1 = "contractor"
elseif txtreport = "PINJAMANPETANIPLASMA" then
	tablename = "plasmaloan"
	jointable1 = "koperasi"
elseif txtreport = "HUTANGKONTRAKTOR" then
	tablename = "apcontractor"
	jointable1 = "contractor"
elseif txtreport = "UANGMUKAKONTRAKTOR" then
	tablename = "contractoradvance"
	jointable1 = "contractor"
elseif txtreport = "HUTANGRETENSI" then
	tablename = "debt_contractretention"
	jointable1 = "contractor"
end if

	strPusingan = " SELECT fcba" & _
	              "        , code" & _
	              "        , description" & _
	              "        , SUM(opening) opening" & _
	              "        , SUM(debet) debet" & _
	              "        , SUM(credit) credit" & _
	              "        , SUM(opening) + SUM(debet) + SUM(credit) ending" & _
	              " FROM (" & _
	              "     SELECT ap.fcba" & _
				  "            , ap.fccode code" & _
				  "            , sp.fcname description" & _
				  "            , TO_CHAR(LAST_DAY(ap.transdate), 'mm/dd/yyyy') transdate" & _
				  "            , ap.opening" & _
				  "            , ap.debet" & _
				  "            , ap.credit" & _
				  "     FROM "&tablename&" ap" & _
				  "     LEFT JOIN "&jointable1&" sp ON (ap.fcba = sp.fcba AND ap.fccode = sp.fccode)" & _
				  "     LEFT JOIN businessunit bu ON (ap.fcba = bu.fccode)" & _
				  "     WHERE ap.period_year = '" & v_txtYear & "'" & _
				  "           AND ap.period_month = '" & v_txtMonth & "'"
	if "" & v_txtFcba <> "" AND "" & v_txtFcba <> "--ALL--" then
		strPusingan = strPusingan & "               AND ap.fcba = '" & v_txtFcba & "'"
	end if
								
	if "" & txtcompany <> "" then
		strPusingan = strPusingan & "               AND bu.fccompanycode = '" & txtcompany & "'"
	end if
	strPusingan = strPusingan & " ) GROUP BY fcba, code, description"
	strPusingan = strPusingan & " HAVING ( NVL(SUM(opening),0) <> 0 OR NVL(SUM(debet),0) <> 0 OR NVL(SUM(credit),0) <> 0 )"
	strPusingan = strPusingan & " ORDER BY fcba, code"
	
	'response.write strPusingan
									
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

Function f_rowspan_unit(data_array__, fcba__)
	result_ = 0
	arr_1 = data_array__
	for indx = 0 to UBound(arr_1, 2)
		if Cstr(arr_1(0, indx)) = Cstr(fcba__) then
			result_ = result_ +1
		end if
	next
	f_rowspan_unit = result_
End Function

Function f_judul(data_array__, where__)
	result_ = ""
	arr_1 = data_array__
	for indx = 0 to UBound(arr_1, 2)
		if Cstr(arr_1(0, indx)) = Cstr(where__) then
			result_ = Cstr(arr_1(1, indx))
		end if
	next
	f_judul = result_
End Function

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
<title>Buku Dua - iPlaS</title>
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
						mon = Month(Session("v_Costbook"))
						If len(v_txtMonth)=0 then v_txtMonth=mon

						Yr = Year(Session("v_Costbook"))
						MinYr = 2015
						MaxYr = Yr  + 2
						If len(v_txtYear)= 0 then v_txtYear = Yr
					%>
							<td width="4">
								<select name="txtYear">
					<%
						For i = MinYr to MaxYr
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
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
							<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
					<%
						Next
					%>
								</select>
							</td>
							<td>
								<font face="verdana" size="1"><%=titleCompany%></font>
							</td>
							<td>
								<select name="txtcompany" >
									<%
									strFcba	=	" SELECT DISTINCT fccompanycode FROM businessunit" & _
												" WHERE user_company_api.getcompanycode('" & Session("CodeName") & "',fccode) IS NOT NULL"
									
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
									<option value="<%=rsFcba("fccompanycode")%>" <%if txtcompany = rsFcba("fccompanycode") then%> Selected <%End If%>><%=rsFcba("fccompanycode")%></option>
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
								<font face="verdana" size="1"><%=titleSite%></font>
							</td>
							<td>
								<select name="txtFcba" >
									<option value="--ALL--" <%if v_txtFcba = "--ALL--" then%> Selected <%End If%>>--ALL--</option>
									<%
									Dim rsFcba, strFcba
									strFcba	=	" SELECT fccode, orderdata FROM businessunit" & _
												" WHERE user_company_api.getcompanycode('" & Session("CodeName") & "',fccode) IS NOT NULL" & _
												" ORDER BY orderdata"
									
									Set rsFcba = Server.CreateObject("ADODB.Recordset")
									rsFcba.Open strFcba, connstr, adOpenStatic

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
								<font face="verdana" size="1">Report</font>
							</td>
							<td>
								<select name="txtreport" >
									<%if not arr_param_eof then%>
										<%for xi=0 to ubound(arr_param,2)%>
											<option value="<%=arr_param(0,xi)%>" <%if txtreport = cstr(arr_param(0,xi)) then%> Selected <%End If%>><%=arr_param(1,xi)%></option>
										<%next%>
									<%end if%>
								</select>
							</td>
							<td><input type="submit" value="View" /></td>
						</tr>				
					</table>
				</div>
			</form>
		</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead style="border-left:1px #000000 solid;">
	<tr>
         <th colspan="7" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
		 <div id="header-wrapper">
			<div id="head_judul">
				<h4><%If Len(Trim(varCompanyName)) > 0 Then%>PT. <%End If%><%=varCompanyName%></h4>
				<h5><%=varFcbaName%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2><%=f_judul(arr_param, txtreport)%></h3>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
			</div>
          </div>  
		 </th>
	</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;"><%=titleSite%></th>
		<th style="background-color:#00A300;color:#fff;">Code</th>
		<th style="background-color:#00A300;color:#fff;">Description</th>
		<th style="background-color:#00A300;color:#fff;">Opening Balance</th>
		<th style="background-color:#00A300;color:#fff;">Debet</th>
		<th style="background-color:#00A300;color:#fff;">Credit</th>
		<th style="background-color:#00A300;color:#fff;">Ending Balance</th>
	</tr>
    </thead>
	<tbody>
<%
Set rsPusingan = Server.CreateObject("ADODB.Recordset")
rsPusingan.Open strPusingan, connstr, adOpenStatic

arr_data_eof = true
if not rsPusingan.eof then
	arr_data = rsPusingan.getrows()
	arr_data_eof = false
end if
rsPusingan.Close
set rsPusingan = Nothing

if not arr_data_eof then
	st_opening = 0
	st_debet = 0
	st_credit = 0
	st_ending = 0
	
	gt_opening = 0
	gt_debet = 0
	gt_credit = 0
	gt_ending = 0
	
	span_fcba = 0
	for xi=0 to ubound(arr_data,2)
		span_fcba = span_fcba +1
		curr_ba = arr_data(0, xi)
		max_span_fcba = f_rowspan_unit(arr_data, arr_data(0, xi))
%>
		<tr>
			<%if curr_ba <> prev_ba then%>
			<td align="center" valign="top" rowspan=<%=max_span_fcba%>><%=arr_data(0, xi)%></td>
			<%end if%>
			<td align="center"><%=arr_data(1, xi)%></td>
			<td align="left"><%=arr_data(2, xi)%></td>
			<td align="right"><%=FormatNumber(arr_data(3, xi), 0)%></td>
			<td align="right"><%=FormatNumber(arr_data(4, xi), 0)%></td>
			<td align="right"><%=FormatNumber(arr_data(5, xi), 0)%></td>
			<td align="right"><%=FormatNumber(arr_data(6, xi), 0)%></td>			
		</tr>
<%
		st_opening = st_opening + cdbl(to_nol(arr_data(3, xi)))
		st_debet = st_debet + cdbl(to_nol(arr_data(4, xi)))
		st_credit = st_credit + cdbl(to_nol(arr_data(5, xi)))
		st_ending = st_ending + cdbl(to_nol(arr_data(6, xi)))
		
		gt_opening = gt_opening + cdbl(to_nol(arr_data(3, xi)))
		gt_debet = gt_debet + cdbl(to_nol(arr_data(4, xi)))
		gt_credit = gt_credit + cdbl(to_nol(arr_data(5, xi)))
		gt_ending = gt_ending + cdbl(to_nol(arr_data(6, xi)))

		prev_ba = curr_ba
		if span_fcba = max_span_fcba then
			response.write "<tr style='background-color:#F4F4A0;'>"
			response.write "	<td align='center' colspan='3'><b>Total "&arr_data(0, xi)&"</b></td>"
			response.write "	<td align='right'><b>"&FormatNumber(st_opening,0)&"</b></td>"
			response.write "	<td align='right'><b>"&FormatNumber(st_debet,0)&"</b></td>"
			response.write "	<td align='right'><b>"&FormatNumber(st_credit,0)&"</b></td>"
			response.write "	<td align='right'><b>"&FormatNumber(st_ending,0)&"</b></td>"
			response.write "</tr>"
		
			st_opening = 0
			st_debet = 0
			st_credit = 0
			st_ending = 0
		
			span_fcba = 0
		end if		
	next
	
	response.write "<tr><td colspan='7'></td></tr>"
	response.write "<tr style='background-color:#FFFF00;'>"
	response.write "	<td align='center' colspan='3'><b>Grand Total</b></td>"
	response.write "	<td align='right'><b>"&FormatNumber(gt_opening,0)&"</b></td>"
	response.write "	<td align='right'><b>"&FormatNumber(gt_debet,0)&"</b></td>"
	response.write "	<td align='right'><b>"&FormatNumber(gt_credit,0)&"</b></td>"
	response.write "	<td align='right'><b>"&FormatNumber(gt_ending,0)&"</b></td>"
	response.write "</tr>"
else
	response.write "<tr><td colspan='7'>NO DATA AVAILABLE</td></tr>"
end if

%>

</tbody>
</table>

   <table width="90%" class="data">
	   <tr>
 		<td width="50%" align="center"><font face="Verdana" size="1">Approved by,</font></td>
		<td width="50%" align="center"><font face="Verdana" size="1">Prepared by,</font></td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <%
		DIM rsttd,ttd
		
		ttd = "SELECT FCMANAGER,FCKTU FROM BusinessUnit WHERE FCCODE = '" & v_txtFcba & "'"
		
	    Set rsttd = Server.CreateObject("ADODB.Recordset")
		rsttd.Open ttd, connstr, adOpenStatic

		If Not rsttd.EOF Then
	   %>
	   <tr> 
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=rsttd("FCMANAGER")%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=rsttd("FCKTU")%></u></strong></td>
	   </tr>
	<%  End If
%> 
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>
   </body>
   </html>
