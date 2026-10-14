<% ValidateSession %>
<!--#include virtual="/conndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
maxCol = 0

txtYear = Request("txtYear")
txtMonth = Request("txtMonth")

txttdate = Request("txttdate")

If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))
If "" & txtMonth = "" Then txtMonth = Month(Session("v_Costbook"))
If "" & txttdate = "" Then txttdate = textDate(date)

txttdate = txtMonth & "" & txtYear

add_col_bud_prod_qty = ""
d_span = 28
	
v_txtFcba = Request("txtFcba")
v_txtDivisi = Request("txtDivisi")

  
    if  v_txtfcba <> "--ALL--"  and v_txtfcba <> "" then
       whereCons = whereCons & " And fcba = '" & v_txtfcba & "' "
	end if
    ' selectCons = selectCons & ",fcba"
    ' orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
	q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & v_txtfcba & "' "
	set rsCat = connStr.Execute(q)
	if not rsCat.eof then
	   fcbaname = rsCat("fcName")
	   CompanyName=rsCat("FCCOMPANYNAME")
	   MgrName = rsCat("FCMANAGER")
	   KTUName = rsCat("FCKTU")
	   CropName=rsCat("FCCROP")
	   FcbaType=rsCat("fcType")
else
    nSpanTotal = nSpanTotal - 1 
end if

	if  v_txtDivisi <> "--ALL--"  and v_txtDivisi <> "" then
    	whereCons = whereCons & "AND INSTR('"& v_txtDivisi &"', SUBSTR(widivision,1,6) ) <> 0 "
	end if

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

arr_field_eof = true
strPusingan	= "	SELECT fcba,widivision, " & _
			  "	SUM(DECODE(BULAN,'01',TOTAL,0))JANUARI_TTL , " & _
			  "	SUM(DECODE(BULAN,'02',TOTAL,0))FEBRUARI_TTL, " & _
		 	  "	SUM(DECODE(BULAN,'03',TOTAL,0))MARET_TTL," & _
			  "	SUM(DECODE(BULAN,'04',TOTAL,0))APRIL_TTL," & _
			  "	SUM(DECODE(BULAN,'05',TOTAL,0))MEI_TTL," & _
			  "	SUM(DECODE(BULAN,'06',TOTAL,0))JUNI_TTL," & _
			  "	SUM(DECODE(BULAN,'07',TOTAL,0))JULI_TTL," & _
			  "	SUM(DECODE(BULAN,'08',TOTAL,0))AGUSTUS_TTL," & _
			  "	SUM(DECODE(BULAN,'09',TOTAL,0))SEPTEMBER_TTL," & _
			  "	SUM(DECODE(BULAN,'10',TOTAL,0))OKTOBER_TTL," & _
			  "	SUM(DECODE(BULAN,'11',TOTAL,0))NOVEMBER_TTL," & _
			  "	SUM(DECODE(BULAN,'12',TOTAL,0))DESEMBER_TTL," & _
		 	  "	SUM(DECODE(BULAN,'01',HARI,0))JANUARI , " & _
			  "	SUM(DECODE(BULAN,'02',HARI,0))FEBRUARI, " & _
			  "	SUM(DECODE(BULAN,'03',HARI,0))MARET," & _
			  "	SUM(DECODE(BULAN,'04',HARI,0))APRIL," & _
			  "	SUM(DECODE(BULAN,'05',HARI,0))MEI," & _
			  "	SUM(DECODE(BULAN,'06',HARI,0))JUNI," & _
			  "	SUM(DECODE(BULAN,'07',HARI,0))JULI," & _
			  "	SUM(DECODE(BULAN,'08',HARI,0))AGUSTUS," & _
			  "	SUM(DECODE(BULAN,'09',HARI,0))SEPTEMBER," & _
			  "	SUM(DECODE(BULAN,'10',HARI,0))OKTOBER," & _
			  "	SUM(DECODE(BULAN,'11',HARI,0))NOVEMBER," & _
			  "	SUM(DECODE(BULAN,'12',HARI,0))DESEMBER" & _
			  "	FROM (" & _
			  "	select TO_CHAR(widate,'YYYY')AS TAHUN,TO_CHAR(widate,'MM')AS BULAN,fcba,widivision,sum(TOTRNINCHES) as total,sum(TOTRNDAYS) as HARI" & _
			  "	from iplasprod.weather" & _
			  "	where to_char(widate,'yyyy')='"& txtYear &"' "& whereCons &""& _
			  "	group by TO_CHAR(widate,'YYYY'),TO_CHAR(widate,'MM'),fcba,widivision" & _
			  "	)" & _
			  "	group by fcba,widivision" & _
			  "	order by fcba,widivision"


' response.write strPusingan
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
' response.write strBaInfo
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
<title>ePlaS Reporting - Curah Hujan Bulanan</title>
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
							<th>KEBUN</th>
							<th>TAHUN</th>
							<th>DIVISI</th>
						</tr>
						<tr>
							<td>
								<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
									<% 
									Set rsFcba = Server.CreateObject("ADODB.Recordset")
									rsFcba.CursorLocation = adUseClient
									strsql =    " SELECT '--ALL--' as fcba from DUAL UNION " & _  
									            " select fccode fcba from businessunit " & _
												" where " & _
												" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL " 
									  
									
									rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
									if not rsFcba.eof then
										DO while not rsFcba.eof
											if v_txtFcba = rsFcba("fcba") then
											%>			
												<option value="<%=rsFcba("fcba")%>" Selected><%=rsFcba("fcba")%></option>
											<% Else %>	
												<option value="<%=rsFcba("fcba")%>"><%=rsFcba("fcba")%></option>
											<% End If
											
											rsFcba.movenext
										loop
									End if
									rsFcba.Close
									set rsFcba = Nothing
									%>
								</select>
							</td> 
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
							<td>
								<select name="txtDivisi"  multiple class="chosen-select">
									<option>--ALL--</option>
									<% 
									Set rsFcba = Server.CreateObject("ADODB.Recordset")
									rsFcba.CursorLocation = adUseClient
									strsql =    " SELECT DISTINCT WIDIVISION FROM WEATHER WHERE FCBA = '"& v_txtFcba &"' ORDER BY WIDIVISION" 
									  
									
									rsFcba.Open strsql, connStr, adOpenForwardOnly, adLockReadOnly	
									if not rsFcba.eof then
										DO while not rsFcba.eof
											if v_txtDivisi = rsFcba("WIDIVISION") then
											%>			
												<option value="<%=rsFcba("WIDIVISION")%>" Selected><%=rsFcba("WIDIVISION")%></option>
											<% Else %>	
												<option value="<%=rsFcba("WIDIVISION")%>"><%=rsFcba("WIDIVISION")%></option>
											<% End If
											
											rsFcba.movenext
										loop
									End if
									rsFcba.Close
									set rsFcba = Nothing
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
			
			<div id="subheader" class="centered">
				<h2>LAPORAN CURAH HUJAN BULANAN</h2>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(txtMonth)%>&nbsp;<%=txtYear%></h4>
			</div>
          </div>  
		 </th>
	</tr>
	 <tr>
					<th rowspan="2">Kebun</th>
					<th rowspan="2">WidDivision</th>
					<th colspan="13">Curah Hujan (mm)</th>
					<th colspan="13">Hari Hujan (Hari)</th>
				</tr>

				<tr>
					<th>Jan</th>
					<th>Feb</th>
					<th>Mar</th>
					<th>Apr</th>
					<th>Mei</th>
					<th>Jun</th>
					<th>Jul</th>
					<th>Aug</th>
					<th>Sep</th>
					<th>Okt</th>
					<th>Nov</th>
					<th>Des</th>
					<th>Total</th>

					<th>Jan</th>
					<th>Feb</th>
					<th>Mar</th>
					<th>Apr</th>
					<th>Mei</th>
					<th>Jun</th>
					<th>Jul</th>
					<th>Aug</th>
					<th>Sep</th>
					<th>Okt</th>
					<th>Nov</th>
					<th>Des</th>
					<th>Total</th>
					
				</tr>

    </thead>
	<tbody>
		<%
		Set rsCat = connStr.Execute(strPusingan)
		
		v_JANUARI_TTL = 0
		v_FEBRUARI_TTL = 0
		v_MARET_TTL = 0 
		v_APRIL_TTL = 0
		v_MEI_TTL = 0
		v_JUNI_TTL = 0
		v_JULI_TTL = 0
		v_AGUSTUS_TTL = 0
		v_SEPTEMBER_TTL = 0
		v_OKTOBER_TTL = 0
		v_NOVEMBER_TTL = 0
		v_DESEMBER_TTL = 0

		v_JANUARI = 0
		v_FEBRUARI = 0
		v_MARET = 0
		v_APRIL = 0
		v_MEI = 0
		v_JUNI = 0
		v_JULI = 0
		v_AGUSTUS = 0
		v_SEPTEMBER = 0
		v_OKTOBER = 0
		v_NOVEMBER = 0
		v_DESEMBER = 0
		

		v_TOTALBULANAN = 0
		v_TOTALHARIAN = 0

		strPusingan	= "	SELECT COUNT(*) TOTAL FROM( SELECT fcba,widivision, " & _
					  "	SUM(DECODE(BULAN,'01',TOTAL,0))JANUARI_TTL , " & _
					  "	SUM(DECODE(BULAN,'02',TOTAL,0))FEBRUARI_TTL, " & _
				 	  "	SUM(DECODE(BULAN,'03',TOTAL,0))MARET_TTL," & _
					  "	SUM(DECODE(BULAN,'04',TOTAL,0))APRIL_TTL," & _
					  "	SUM(DECODE(BULAN,'05',TOTAL,0))MEI_TTL," & _
					  "	SUM(DECODE(BULAN,'06',TOTAL,0))JUNI_TTL," & _
					  "	SUM(DECODE(BULAN,'07',TOTAL,0))JULI_TTL," & _
					  "	SUM(DECODE(BULAN,'08',TOTAL,0))AGUSTUS_TTL," & _
					  "	SUM(DECODE(BULAN,'09',TOTAL,0))SEPTEMBER_TTL," & _
					  "	SUM(DECODE(BULAN,'10',TOTAL,0))OKTOBER_TTL," & _
					  "	SUM(DECODE(BULAN,'11',TOTAL,0))NOVEMBER_TTL," & _
					  "	SUM(DECODE(BULAN,'12',TOTAL,0))DESEMBER_TTL," & _
				 	  "	SUM(DECODE(BULAN,'01',HARI,0))JANUARI , " & _
					  "	SUM(DECODE(BULAN,'02',HARI,0))FEBRUARI, " & _
					  "	SUM(DECODE(BULAN,'03',HARI,0))MARET," & _
					  "	SUM(DECODE(BULAN,'04',HARI,0))APRIL," & _
					  "	SUM(DECODE(BULAN,'05',HARI,0))MEI," & _
					  "	SUM(DECODE(BULAN,'06',HARI,0))JUNI," & _
					  "	SUM(DECODE(BULAN,'07',HARI,0))JULI," & _
					  "	SUM(DECODE(BULAN,'08',HARI,0))AGUSTUS," & _
					  "	SUM(DECODE(BULAN,'09',HARI,0))SEPTEMBER," & _
					  "	SUM(DECODE(BULAN,'10',HARI,0))OKTOBER," & _
					  "	SUM(DECODE(BULAN,'11',HARI,0))NOVEMBER," & _
					  "	SUM(DECODE(BULAN,'12',HARI,0))DESEMBER" & _
					  "	FROM (" & _
					  "	select TO_CHAR(widate,'YYYY')AS TAHUN,TO_CHAR(widate,'MM')AS BULAN,fcba,widivision,sum(TOTRNINCHES) as total,sum(TOTRNDAYS) as HARI" & _
					  "	from iplasprod.weather" & _
					  "	where to_char(widate,'yyyy')='"& txtYear &"' "& whereCons &""& _
					  "	group by TO_CHAR(widate,'YYYY'),TO_CHAR(widate,'MM'),fcba,widivision" & _
					  "	)" & _
					  "	group by fcba,widivision" & _
					  "	order by fcba,widivision) "


		' response.write strPusingan
		Set qORS = Server.CreateObject("ADODB.Recordset")
		qORS.Open strPusingan, connstr, adOpenStatic
		If Not qORS.EOF Then
			Total = cdbl(to_nol(qORS("TOTAL")))
		End If
		qORS.Close
		set qORS = Nothing

		if Not rsCat.eof then
			While Not rsCat.EOF
			TOTALBULANAN = cdbl(to_nol(rsCat("JANUARI_TTL"))) + cdbl(to_nol(rsCat("FEBRUARI_TTL"))) + cdbl(to_nol(rsCat("MARET_TTL"))) + cdbl(to_nol(rsCat("APRIL_TTL"))) + cdbl(to_nol(rsCat("MEI_TTL"))) + cdbl(to_nol(rsCat("JUNI_TTL"))) + cdbl(to_nol(rsCat("JULI_TTL"))) + cdbl(to_nol(rsCat("AGUSTUS_TTL"))) + cdbl(to_nol(rsCat("SEPTEMBER_TTL"))) + cdbl(to_nol(rsCat("OKTOBER_TTL"))) + cdbl(to_nol(rsCat("NOVEMBER_TTL"))) + cdbl(to_nol(rsCat("DESEMBER_TTL")))

			TOTALHARIAN = cdbl(to_nol(rsCat("JANUARI"))) + cdbl(to_nol(rsCat("FEBRUARI"))) + cdbl(to_nol(rsCat("MARET"))) + cdbl(to_nol(rsCat("APRIL"))) + cdbl(to_nol(rsCat("MEI"))) + cdbl(to_nol(rsCat("JUNI"))) + cdbl(to_nol(rsCat("JULI"))) + cdbl(to_nol(rsCat("AGUSTUS"))) + cdbl(to_nol(rsCat("SEPTEMBER"))) + cdbl(to_nol(rsCat("OKTOBER"))) + cdbl(to_nol(rsCat("NOVEMBER"))) + cdbl(to_nol(rsCat("DESEMBER"))) 
		%>
		<tr>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcba")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("widivision")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JANUARI_TTL")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FEBRUARI_TTL")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("MARET_TTL")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("APRIL_TTL")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("MEI_TTL")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JUNI_TTL")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JULI_TTL")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("AGUSTUS_TTL")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SEPTEMBER_TTL")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("OKTOBER_TTL")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("NOVEMBER_TTL")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DESEMBER_TTL")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALBULANAN%></td>

			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JANUARI")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("FEBRUARI")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("MARET")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("APRIL")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("MEI")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JUNI")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JULI")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("AGUSTUS")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("SEPTEMBER")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("OKTOBER")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("NOVEMBER")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("DESEMBER")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALHARIAN%></td>
		</tr>
	<% 
		v_JANUARI_TTL = v_JANUARI_TTL +  cdbl(to_nol(rsCat("JANUARI_TTL")))  / Total
		v_FEBRUARI_TTL = v_FEBRUARI_TTL +  cdbl(to_nol(rsCat("FEBRUARI_TTL")))  / Total
		v_MARET_TTL = v_MARET_TTL +  cdbl(to_nol(rsCat("MARET_TTL")))   / Total
		v_APRIL_TTL = v_APRIL_TTL +  cdbl(to_nol(rsCat("APRIL_TTL")))  / Total
		v_MEI_TTL = v_MEI_TTL +  cdbl(to_nol(rsCat("MEI_TTL")))  / Total
		v_JUNI_TTL = v_JUNI_TTL +  cdbl(to_nol(rsCat("JUNI_TTL")))  / Total
		v_JULI_TTL = v_JULI_TTL +  cdbl(to_nol(rsCat("JULI_TTL"))) / Total
		v_AGUSTUS_TTL = v_AGUSTUS_TTL +  cdbl(to_nol(rsCat("AGUSTUS_TTL"))) / Total 
		v_SEPTEMBER_TTL = v_SEPTEMBER_TTL +  cdbl(to_nol(rsCat("SEPTEMBER_TTL"))) / Total
		v_OKTOBER_TTL = v_OKTOBER_TTL +  cdbl(to_nol(rsCat("OKTOBER_TTL"))) / Total
		v_NOVEMBER_TTL = v_NOVEMBER_TTL +  cdbl(to_nol(rsCat("NOVEMBER_TTL"))) / Total
		v_DESEMBER_TTL = v_DESEMBER_TTL +  cdbl(to_nol(rsCat("DESEMBER_TTL"))) / Total

		v_JANUARI = v_JANUARI +  cdbl(to_nol(rsCat("JANUARI")))   / Total
		v_FEBRUARI  = v_FEBRUARI +  cdbl(to_nol(rsCat("FEBRUARI")))  / Total
		v_MARET = v_MARET +  cdbl(to_nol(rsCat("MARET"))) / Total
		v_APRIL = v_APRIL +  cdbl(to_nol(rsCat("APRIL"))) / Total
		v_MEI = v_MEI +  cdbl(to_nol(rsCat("MEI"))) / Total
		v_JUNI = v_JUNI +  cdbl(to_nol(rsCat("JUNI"))) / Total
		v_JULI = v_JULI +  cdbl(to_nol(rsCat("JULI"))) / Total
		v_AGUSTUS = v_AGUSTUS +  cdbl(to_nol(rsCat("AGUSTUS"))) / Total 
		v_SEPTEMBER = v_SEPTEMBER +  cdbl(to_nol(rsCat("SEPTEMBER"))) / Total
		v_OKTOBER = v_OKTOBER +  cdbl(to_nol(rsCat("OKTOBER"))) / Total
		v_NOVEMBER = v_NOVEMBER +  cdbl(to_nol(rsCat("NOVEMBER"))) / Total
		v_DESEMBER = v_DESEMBER +  cdbl(to_nol(rsCat("DESEMBER"))) / Total

		v_TOTALBULANAN = v_TOTALBULANAN +  cdbl(to_nol(TOTALBULANAN)) / Total
		v_TOTALHARIAN = v_TOTALHARIAN +  cdbl(to_nol(TOTALHARIAN)) / Total


		rsCat.MoveNext
		Wend
	Else
		response.write "<tr><td colspan='26'>No Data Available</td></tr>"
	End If
	rsCat.close()
	rsCat=null%>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
		<tr bgcolor="#F4F4A0">
			<td align="center" colspan="2"><strong>AVG</strong></td>
			<td align="right"><strong><%=FormatNumber(v_JANUARI_TTL,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_FEBRUARI_TTL,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_MARET_TTL,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_APRIL_TTL,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_MEI_TTL,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_JUNI_TTL,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_JULI_TTL,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_AGUSTUS_TTL,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_SEPTEMBER_TTL,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_OKTOBER_TTL,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_NOVEMBER_TTL,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_DESEMBER_TTL,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_TOTALBULANAN,0)%></strong></td>
 
			<td align="right"><strong><%=FormatNumber(v_JANUARI,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_FEBRUARI,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_MARET,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_APRIL,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_MEI,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_JUNI,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_JULI,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_AGUSTUS,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_SEPTEMBER,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_OKTOBER,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_NOVEMBER,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_DESEMBER,0)%></strong></td>
			<td align="right"><strong><%=FormatNumber(v_TOTALHARIAN,0)%></strong></td>
		</tr>
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