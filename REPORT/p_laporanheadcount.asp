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
d_span = 26
	
v_txtFcba = Request("txtFcba")
v_txtGangCode = Request("txtGangCode")
v_txtDivision = Request("txtDivision")
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
  
				if  v_txtfcba <> "--ALL--"  and v_txtfcba <> "" then
			       whereCons = whereCons & " And fcba = '" & v_txtfcba & "'"
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

				' if v_txtDivision <> "--SUMMARY--"  And v_txtDivision <> "" then
			 '      if v_txtDivision <>  "--ALL--"  then
				' 	whereCons = whereCons & " AND SECTIONNAME = '" & v_txtDivision & "'"
			 '      end if
    '               selectCons = selectCons &  " , SECTIONNAME "
				'   orderCons  = orderCons & ",SECTIONNAME "
			 '   else
			 '   nSpanTotal = nSpanTotal - 1  
			 '   end if

			 '     if  v_txtGangCode <> "--ALL--"  and v_txtGangCode <> "" then
				'        whereCons = whereCons & " And fccode = '" & v_txtGangCode & "' "
				' 	end if
				'     ' selectCons = selectCons & ",fcba"
				'     ' orderCons  = orderCons & ",BUSINESSUNIT_API.get_OrderData(fcba)"
				' 	q = "Select fccode from gang where fcCode='" & v_txtGangCode & "'"
				' 	set rsCat = connStr.Execute(q)
				' 	if not rsCat.eof then
				' 	   fcCode = rsCat("fcCode")
				' else
				'     nSpanTotal = nSpanTotal - 1 
				' end if

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

arr_field_eof = true
strPusingan	= 	" select sectionname, fcba" & _
				" ,SUM(decode(emp.presentgrade,'KHT',Jumlah_TK_PNN,0)) as JMLKHT_TK_PNN" & _
				" ,SUM(decode(emp.presentgrade,'KHT',Jumlah_TK_MD,0)) as JMLKHT_TK_MD" & _
				" ,SUM(decode(emp.presentgrade,'KHT',Jumlah_TK_PML,0)) as JMLKHT_TK_PML" & _
				" ,SUM(decode(emp.presentgrade,'KHT',Jumlah_TK_TR,0)) as JMLKHT_TK_TR" & _
				" ,SUM(decode(emp.presentgrade,'KHT',Jumlah_TK_KU,0)) as JMLKHT_TK_KU" & _
				" ,SUM(decode(emp.presentgrade,'KHL',Jumlah_TK_PNN,0)) as JMLKHL_TK_PNN" & _
				" ,SUM(decode(emp.presentgrade,'KHL',Jumlah_TK_MD,0)) as JMLKHL_TK_MD" & _
				" ,SUM(decode(emp.presentgrade,'KHL',Jumlah_TK_PML,0)) as JMLKHL_TK_PML" & _
				" ,SUM(decode(emp.presentgrade,'KHL',Jumlah_TK_TR,0)) as JMLKHL_TK_TR" & _
				" ,SUM(decode(emp.presentgrade,'KHL',Jumlah_TK_KU,0)) as JMLKHL_TK_KU" & _
				" ,SUM(decode(emp.presentgrade,'KBT',Jumlah_TK_PNN,0)) as JMLKBT_TK_PNN" & _
				" ,SUM(decode(emp.presentgrade,'KBT',Jumlah_TK_MD,0)) as JMLKBT_TK_MD" & _
				" ,SUM(decode(emp.presentgrade,'KBT',Jumlah_TK_PML,0)) as JMLKBT_TK_PML" & _
				" ,SUM(decode(emp.presentgrade,'KBT',Jumlah_TK_TR,0)) as JMLKBT_TK_TR" & _
				" ,SUM(decode(emp.presentgrade,'KBT',Jumlah_TK_KU,0)) as JMLKBT_TK_KU" & _
				" from (" & _
				" SELECT fcba," & _
				"       sectionname," & _
				"       presentgrade," & _
				"       COUNT (presentgrade) AS Jumlah_TK_PNN," & _
				"       0 Jumlah_TK_MD," & _
				"       0 Jumlah_TK_TR," & _
				"       0 Jumlah_TK_KU," & _
				"       0 Jumlah_TK_PML" & _
				"  FROM iplasprod.employee_hist_tab" & _
				" WHERE  dateterminate IS NULL" & _
				"       AND SUBSTR (gangcode, 1, 2) IN ('PN')" & _
				"       AND year = '"& txtyear &"'" & _
				"       AND month = '"& txtMonth &"'" & _
				"		"& whereCons &""& _
				" GROUP BY fcba," & _
				"       sectionname," & _
				"       presentgrade" & _
				" UNION ALL" & _
				" SELECT fcba," & _
				"       sectionname," & _
				"       presentgrade," & _
				"      0 Jumlah_TK_PNN," & _
				"       COUNT (presentgrade) AS Jumlah_TK_MD," & _
				"       0 Jumlah_TK_TR," & _
				"       0 Jumlah_TK_KU," & _
				"       0 Jumlah_TK_PML" & _
				"  FROM iplasprod.employee_hist_tab" & _
				" WHERE dateterminate IS NULL" & _
				"       AND SUBSTR (gangcode, 1, 2) IN ('MD')" & _
				"       AND year = '"& txtyear &"'" & _
				"       AND month = '"& txtMonth &"'" & _
				"		"& whereCons &""& _
				" GROUP BY fcba," & _
				"       sectionname," & _
				"       presentgrade" & _
				" UNION ALL" & _
				" SELECT fcba," & _
				"       sectionname," & _
				"       presentgrade," & _
				"       0 Jumlah_TK_PNN," & _
				"       0 Jumlah_TK_MD," & _
				"       COUNT (presentgrade) AS Jumlah_TK_TR," & _
				"       0 Jumlah_TK_KU," & _
				"       0 Jumlah_TK_PML" & _
				"  FROM iplasprod.employee_hist_tab" & _
				" WHERE dateterminate IS NULL" & _
				"       AND SUBSTR (gangcode, 1, 2) IN ('OS', 'MT', 'TB')" & _
				"       AND year = '"& txtyear &"'" & _
				"       AND month = '"& txtMonth &"'" & _
				"		"& whereCons &""& _
				" GROUP BY fcba," & _
				"       sectionname," & _
				"       presentgrade" & _
				" UNION ALL" & _
				" SELECT fcba," & _
				"       sectionname," & _
				"       presentgrade," & _
				"       0 Jumlah_TK_PNN," & _
				"       0 Jumlah_TK_MD," & _
				"       0 Jumlah_TK_TR," & _
				"       COUNT (presentgrade) AS Jumlah_TK_KU," & _
				"       0 Jumlah_TK_PML" & _
				"  FROM iplasprod.employee_hist_tab" & _
				" WHERE dateterminate IS NULL" & _
				"       AND SUBSTR (gangcode, 1, 2) IN ('KT', 'KM', 'UM')" & _
				"       AND year = '"& txtyear &"'" & _
				"       AND month = '"& txtMonth &"'" & _
				"		"& whereCons &""& _
				" GROUP BY fcba," & _
				"       sectionname," & _
				"       presentgrade" & _
				" UNION ALL" & _
				" SELECT fcba," & _
				"       sectionname," & _
				"       presentgrade," & _
				"       0 Jumlah_TK_PNN," & _
				"       0 Jumlah_TK_MD," & _
				"       0 Jumlah_TK_TR," & _
				"       0 Jumlah_TK_KU," & _
				"       COUNT (presentgrade) AS Jumlah_TK_PML" & _
				"  FROM iplasprod.employee_hist_tab" & _
				" WHERE dateterminate IS NULL" & _
				"       AND SUBSTR (gangcode, 1, 2) NOT IN ('PN','KT','KM','UM','OS','MT','TB','MD')" & _
				"       AND year = '"& txtyear &"'" & _
				"       AND month = '"& txtMonth &"'" & _
				"		"& whereCons &""& _
				" GROUP BY fcba," & _
				"       sectionname," & _
				"       presentgrade )emp" & _
				"       GROUP BY sectionname, fcba" & _
				"       order by sectionname, fcba" 


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
<title>Headcount Tenaga Kerja - iPlaS Reporting</title>
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
							
							<th>TAHUN</th>
							<th>BULAN</th>
							<th>KEBUN</th>
							<th>&nbsp;</th>
						</tr>
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
								
								<td width="2">
									<select name="txtMonth">
						<% 
							For i = 1 to 12
								
								if i <= 9 then 
									iBulan = 0&i
								else
									iBulan = i
								end if
								i	= iBulan
						%>
								<Option Value="<%=i%>" <%If cint(i) = cint(txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
						<%
							Next
						%>
									</select>
								</td>
								<td>
							<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
							<%
							Dim rsFcba
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
						<td align="center"><input type="submit" value="View" /></td>
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
				<h2>LAPORAN HEADCOUNT TK</h2>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(txtMonth)%>&nbsp;<%=txtYear%></h4>
			</div>
          </div>  
		 </th>
	</tr>
	 <tr>
					<th rowspan="2">Kebun</th>
					<th rowspan="2">Division</th>
					<th colspan="4">Pemeliharaan</th>
					<th colspan="4">Panen</th>
					<th colspan="4">Supervisi</th>
					<th colspan="4">Traksi</th>
					<th colspan="4">Kantor/Umum</th>
					<th colspan="4">Total</th>
				</tr>

				<tr>
					<th>KBT</th>
					<th>KHT</th>
					<th>KHL</th>
					<th>TOTAL</th>

					<th>KBT</th>
					<th>KHT</th>
					<th>KHL</th>
					<th>TOTAL</th>

					<th>KBT</th>
					<th>KHT</th>
					<th>KHL</th>
					<th>TOTAL</th>

					<th>KBT</th>
					<th>KHT</th>
					<th>KHL</th>
					<th>TOTAL</th>

					<th>KBT</th>
					<th>KHT</th>
					<th>KHL</th>
					<th>TOTAL</th>

					<th>KBT</th>
					<th>KHT</th>
					<th>KHL</th>
					<th>TOTAL</th>
				</tr>

    </thead>
	<tbody>
		<%
		Set rsCat = connStr.Execute(strPusingan)
		if Not rsCat.eof then
			While Not rsCat.EOF
			TOTALPML = cdbl(to_nol(rsCat("JMLKHT_TK_PML"))) + cdbl(to_nol(rsCat("JMLKHL_TK_PML"))) + cdbl(to_nol(rsCat("JMLKBT_TK_PML")))
			TOTALPNN = cdbl(to_nol(rsCat("JMLKHT_TK_PNN"))) + cdbl(to_nol(rsCat("JMLKHL_TK_PNN"))) + cdbl(to_nol(rsCat("JMLKBT_TK_PNN")))
			TOTALMD = cdbl(to_nol(rsCat("JMLKHT_TK_MD"))) + cdbl(to_nol(rsCat("JMLKHL_TK_MD"))) + cdbl(to_nol(rsCat("JMLKBT_TK_MD")))
			TOTALKU = cdbl(to_nol(rsCat("JMLKBT_TK_KU"))) + cdbl(to_nol(rsCat("JMLKHT_TK_KU"))) + cdbl(to_nol(rsCat("JMLKHL_TK_KU")))
			TOTALTR = cdbl(to_nol(rsCat("JMLKBT_TK_TR"))) + cdbl(to_nol(rsCat("JMLKHT_TK_TR"))) + cdbl(to_nol(rsCat("JMLKHL_TK_TR")))

			TOTALKBT = cdbl(to_nol(rsCat("JMLKBT_TK_PML"))) + cdbl(to_nol(rsCat("JMLKBT_TK_PNN"))) + cdbl(to_nol(rsCat("JMLKBT_TK_MD"))) +  cdbl(to_nol(rsCat("JMLKBT_TK_KU"))) + cdbl(to_nol(rsCat("JMLKBT_TK_TR")))
			TOTALKHT = cdbl(to_nol(rsCat("JMLKHT_TK_PML"))) + cdbl(to_nol(rsCat("JMLKHT_TK_PNN"))) + cdbl(to_nol(rsCat("JMLKHT_TK_MD"))) +  cdbl(to_nol(rsCat("JMLKHT_TK_KU"))) + cdbl(to_nol(rsCat("JMLKHT_TK_TR")))
			TOTALKHL = cdbl(to_nol(rsCat("JMLKHL_TK_PML"))) + cdbl(to_nol(rsCat("JMLKHL_TK_PNN"))) + cdbl(to_nol(rsCat("JMLKHL_TK_MD"))) +  cdbl(to_nol(rsCat("JMLKHL_TK_KU"))) + cdbl(to_nol(rsCat("JMLKHL_TK_TR")))
			TOTAL = TOTALKBT + TOTALKHT  + TOTALKHL
			
			TOTALKBTPML = TOTALKBTPML + cdbl(to_nol(rsCat("JMLKBT_TK_PML")))
			TOTALKHTPML = TOTALKHTPML + cdbl(to_nol(rsCat("JMLKHT_TK_PML")))
			TOTALKHLPML = TOTALKHLPML + cdbl(to_nol(rsCat("JMLKHL_TK_PML")))
			TOTALTOTALPML = TOTALTOTALPML + TOTALPML

			TOTALKBTPNN = TOTALKBTPNN + cdbl(to_nol(rsCat("JMLKBT_TK_PNN")))
			TOTALKHTPNN = TOTALKHTPNN + cdbl(to_nol(rsCat("JMLKHT_TK_PNN")))
			TOTALKHLPNN = TOTALKHLPNN + cdbl(to_nol(rsCat("JMLKHL_TK_PNN")))
			TOTALTOTALPNN = TOTALTOTALPNN + TOTALPNN

			TOTALKBTMD = TOTALKBTMD + cdbl(to_nol(rsCat("JMLKBT_TK_MD")))
			TOTALKHTMD = TOTALKHTMD + cdbl(to_nol(rsCat("JMLKHT_TK_MD")))
			TOTALKHLMD = TOTALKHLMD + cdbl(to_nol(rsCat("JMLKHL_TK_MD")))
			TOTALTOTALMD = TOTALTOTALMD + TOTALMD

			TOTALKBTTR = TOTALKBTTR + cdbl(to_nol(rsCat("JMLKBT_TK_TR")))
			TOTALKHTTR = TOTALKHTTR + cdbl(to_nol(rsCat("JMLKHT_TK_TR")))
			TOTALKHLTR = TOTALKHLTR + cdbl(to_nol(rsCat("JMLKHL_TK_TR")))
			TOTALTOTALTR = TOTALTOTALTR + TOTALTR

			TOTALKBTKU = TOTALKBTKU + cdbl(to_nol(rsCat("JMLKBT_TK_KU")))
			TOTALKHTKU = TOTALKHTKU + cdbl(to_nol(rsCat("JMLKHT_TK_KU")))
			TOTALKHLKU = TOTALKHLKU + cdbl(to_nol(rsCat("JMLKHL_TK_KU")))
			TOTALTOTALKU = TOTALTOTALKU + TOTALKU

			TOTALSKBT = TOTALSKBT + TOTALKBT
			TOTALSKHT = TOTALSKHT + TOTALKHT
			TOTALSKHL = TOTALSKHL + TOTALKHL
			TOTALSTOTAL = TOTALSTOTAL + TOTAL
		%>
		<tr>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fcba")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("sectionname")%></td>

			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JMLKBT_TK_PML")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JMLKHT_TK_PML")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JMLKHL_TK_PML")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALPML%></td>

			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JMLKBT_TK_PNN")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JMLKHT_TK_PNN")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JMLKHL_TK_PNN")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALPNN%></td>

			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JMLKBT_TK_MD")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JMLKHT_TK_MD")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JMLKHL_TK_MD")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALMD%></td>

			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JMLKBT_TK_TR")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JMLKHT_TK_TR")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JMLKHL_TK_TR")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALTR%></td>

			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JMLKBT_TK_KU")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JMLKHT_TK_KU")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("JMLKHL_TK_KU")%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALKU%></td>

			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALKBT%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALKHT%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTALKHL%></td>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=TOTAL%></td>			
		</tr>

	<%
		rsCat.MoveNext
		Wend
	Else
		response.write "<tr><td colspan='27'>No Data Available</td></tr>"
	End If
	rsCat.close()
	rsCat=null%>
		<tr>
			<td colspan="2" align="center" style="background: #F4F4A0;"> TOTAL</td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALKBTPML%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALKHTPML%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALKHLPML%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALTOTALPML%></td>

			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALKBTPNN%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALKHTPNN%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALKHLPNN%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALTOTALPNN%></td>

			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALKBTMD%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALKHTMD%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALKHLMD%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALTOTALMD%></td>

			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALKBTKU%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALKHTKU%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALKHLKU%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALTOTALKU%></td>

			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALKBTTR%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALKHTTR%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALKHLTR%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALTOTALTR%></td>

			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALSKBT%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALSKHT%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALSKHL%></td>
			<td align="center" style="background: #F4F4A0;"><font face="Verdana" size="1"><%=TOTALSTOTAL%></td>
		</tr>
	<input  type="hidden" value="<%=txtTable%>" name="txtTable">
	<input  type="hidden" value="<%=txtModule%>" name="txtModule">
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