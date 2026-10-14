<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<% ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

	If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))
	If "" & txtMonth = "" Then txtMonth = Month(Session("v_Costbook"))
	If "" & txttdate = "" Then txttdate = textDate(date)
	v_txtFcba = Request("txtFcba")
	v_txtFcba2 = Request("txtFcba2")

	txttdate = txtMonth & "" & txtYear
	v_txtPOCode		=	trim(request("txtPOCode"))
	compare1 = trim(request("compare1"))

	varDefaultBa = Session("fcBA")

	v_txtsectionname = request("txtsectionname")
	v_txtvehiclecode = request("txtvehiclecode")
	v_txtdatefrom = Request("txtdatefrom")
	v_txtdateuntil = Request("txtdateuntil")
	v_txtgangcode = Request("txtgangcode")
	v_txtempcode = Request("txtempcode")
	If (v_txtdatefrom = "" Or Len(Trim(v_txtdatefrom)) = 0) Then v_txtdatefrom = textdate(date)
	If (v_txtdateuntil = "" Or Len(Trim(v_txtdateuntil)) = 0) Then v_txtdateuntil = textdate(date)

	If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
		v_txtFcba = varDefaultBa
	End If	
		
   if  v_txtfcba <> "" then
       whereCons = whereCons & "  AND fp.fcba = '" & v_txtfcba & "'"
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
	
	strGrnInv		=	"   SELECT fcba," & _
						"          fddate," & _
						"          employeecode," & _
						"          employee_api.get_employeename (fcba, employeecode) employeename," & _
						"          employee_api.get_sectionname (fcba, employeecode) division," & _
						"		   employee_api.get_gangcode (fcba, employeecode) gangcode," & _
						"          total_tbs," & _
						"          kg_tbs_extra_harian," & _
						"          premi_lebih_borong_tbs_harian," & _
						"          total_bucket," & _
						"          kg_brondolan_harian," & _
						"          premi_brondolan_harian," & _
						"          total_premi_kontanan" & _
						"     FROM (  SELECT premi_kontanan.fcba," & _
						"                    premi_kontanan.fddate," & _
						"                    premi_kontanan.employeecode," & _
						"                    ROUND (SUM (premi_kontanan.total_tbs), 0) AS total_tbs," & _
						"                    ROUND (SUM (premi_kontanan.kg_tbs_extra_harian), 0)" & _
						"                       AS kg_tbs_extra_harian," & _
						"                    ROUND (SUM (premi_kontanan.premi_lebih_borong_tbs_harian), 2)" & _
						"                       AS premi_lebih_borong_tbs_harian," & _
						"                    ROUND (SUM (premi_kontanan.total_bucket), 0) AS total_bucket," & _
						"                    ROUND (SUM (premi_kontanan.kg_brondolan_harian), 0)" & _
						"                       AS kg_brondolan_harian," & _
						"                    ROUND (SUM (premi_kontanan.premi_brondolan_harian), 2)" & _
						"                       AS premi_brondolan_harian," & _
						"                    ROUND (" & _
						"                       NVL (" & _
						"                          (  SUM (premi_kontanan.premi_lebih_borong_tbs_harian)" & _
						"                           + SUM (premi_kontanan.premi_brondolan_harian))," & _
						"                          0)," & _
						"                       2)" & _
						"                       AS total_premi_kontanan" & _
						"               FROM (SELECT bkm.fcba," & _
						"                            bkm.fddate," & _
						"                            bkm.employeecode," & _
						"                            hdb.division AS division_blok," & _
						"                            bkm.locationcode," & _
						"                            NVL (bkm.total_tbs, 0) AS total_tbs," & _
						"                            NVL (hdb.kg_tbs_extra_harian, 0)" & _
						"                               AS kg_tbs_extra_harian," & _
						"                            NVL (hdb.premi_lebih_borong_tbs_harian, 0)" & _
						"                               AS premi_lebih_borong_tbs_harian," & _
						"                            NVL (bkm.total_bucket, 0) AS total_bucket," & _
						"                            NVL (hmwn_daily.bbr, 0) AS bbr," & _
						"                            NVL ( (bkm.total_bucket * hmwn_daily.bbr), 0)" & _
						"                               AS kg_brondolan_harian," & _
						"                            NVL (mhr.rate_brondolan, 0) AS rate_brondolan," & _
						"                            NVL (" & _
						"                               (  NVL ( (bkm.total_bucket * hmwn_daily.bbr), 0)" & _
						"                                * NVL (mhr.rate_brondolan, 0))," & _
						"                               0)" & _
						"                               AS premi_brondolan_harian" & _
						"                       FROM (  SELECT iplasprod.attendance_gad.fcba," & _
						"                                      iplasprod.attendance_gad.fddate," & _
						"                                      iplasprod.attendance_gad.employeecode," & _
						"                                      iplasprod.attendance_gad.locationcode," & _
						"                                      SUM (iplasprod.attendance_gad.unit)" & _
						"                                         AS total_tbs," & _
						"                                      SUM (iplasprod.attendance_gad.bucket)" & _
						"                                         AS total_bucket" & _
						"                                 FROM iplasprod.attendance_gad" & _
						"                                WHERE     iplasprod.attendance_gad.jobcode IN" & _
						"                                             ('6502101')" & _
						"                                      AND iplasprod.attendance_gad.attendance = 'KK'" & _
						"                             GROUP BY iplasprod.attendance_gad.fcba," & _
						"                                      iplasprod.attendance_gad.fddate," & _
						"                                      iplasprod.attendance_gad.employeecode," & _
						"                                      iplasprod.attendance_gad.locationcode) bkm" & _
						"                            INNER JOIN" & _
						"                            (  SELECT iplasprod.harvesting_daily_bjr.fcba," & _
						"                                      iplasprod.harvesting_daily_bjr.fddate," & _
						"                                      iplasprod.harvesting_daily_bjr.employeecode," & _
						"                                      iplasprod.harvesting_daily_bjr.division," & _
						"                                      iplasprod.harvesting_daily_bjr.plantingyear," & _
						"                                      iplasprod.harvesting_daily_bjr.presentgrade," & _
						"                                      iplasprod.harvesting_daily_bjr.attendance," & _
						"                                      iplasprod.harvesting_daily_bjr.fieldcode," & _
						"                                      SUM (" & _
						"                                         iplasprod.harvesting_daily_bjr.extraquantity_bydaily)" & _
						"                                         AS kg_tbs_extra_harian," & _
						"                                      SUM (" & _
						"                                         iplasprod.harvesting_daily_bjr.extraamount_bydaily)" & _
						"                                         AS premi_lebih_borong_tbs_harian" & _
						"                                 FROM iplasprod.harvesting_daily_bjr" & _
						"                                WHERE     iplasprod.harvesting_daily_bjr.fcba =" & _
						"                                             'KNS'" & _
						"                                      AND iplasprod.harvesting_daily_bjr.attendance =" & _
						"                                             'KK'" & _
						"                             GROUP BY iplasprod.harvesting_daily_bjr.fcba," & _
						"                                      iplasprod.harvesting_daily_bjr.fddate," & _
						"                                      iplasprod.harvesting_daily_bjr.employeecode," & _
						"                                      iplasprod.harvesting_daily_bjr.division," & _
						"                                      iplasprod.harvesting_daily_bjr.plantingyear," & _
						"                                      iplasprod.harvesting_daily_bjr.presentgrade," & _
						"                                      iplasprod.harvesting_daily_bjr.attendance," & _
						"                                      iplasprod.harvesting_daily_bjr.fieldcode)" & _
						"                            hdb" & _
						"                               ON (    bkm.fcba = hdb.fcba" & _
						"                                   AND TO_NUMBER (TO_CHAR (bkm.fddate, 'mmyyyy')) =" & _
						"                                          TO_NUMBER (" & _
						"                                             TO_CHAR (" & _
						"                                                TO_DATE (hdb.fddate, 'mm/dd/yyyy')," & _
						"                                                'mmyyyy'))" & _
						"                                   AND bkm.employeecode = hdb.employeecode" & _
						"                                   AND bkm.locationcode = hdb.fieldcode)" & _
						"                            INNER JOIN" & _
						"                            (SELECT iplasprod.harvesting_mill_weight_netto.fcba_estate," & _
						"                                    iplasprod.harvesting_mill_weight_netto.harvestdate," & _
						"                                    iplasprod.harvesting_mill_weight_netto.fieldcode," & _
						"                                    iplasprod.harvesting_mill_weight_netto.bbr" & _
						"                               FROM iplasprod.harvesting_mill_weight_netto)" & _
						"                            hmwn_daily" & _
						"                               ON (    bkm.fcba = hmwn_daily.fcba_estate" & _
						"                                   AND bkm.fddate = hmwn_daily.harvestdate" & _
						"                                   AND bkm.locationcode = hmwn_daily.fieldcode)" & _
						"                            INNER JOIN" & _
						"                            (SELECT distinct iplasprod.multilevel_harvesting_rate.fcba," & _
						"                                    iplasprod.multilevel_harvesting_rate.division," & _
						"                                    iplasprod.multilevel_harvesting_rate.planting_year," & _
						"                                    iplasprod.multilevel_harvesting_rate.employeetype," & _
						"                                    iplasprod.multilevel_harvesting_rate.attendance," & _
						"                                    iplasprod.multilevel_harvesting_rate.rate_brondolan" & _
						"                               FROM iplasprod.multilevel_harvesting_rate" & _
						"                              WHERE iplasprod.multilevel_harvesting_rate.attendance =" & _
						"                                       'KK') mhr" & _
						"                               ON (    hdb.fcba = mhr.fcba" & _
						"                                   AND hdb.division = mhr.division" & _
						"                                   AND hdb.plantingyear = mhr.planting_year" & _
						"                                   AND hdb.presentgrade = mhr.employeetype" & _
						"                                   AND hdb.attendance = mhr.attendance))" & _
						"                    premi_kontanan" & _
						"           GROUP BY premi_kontanan.fcba," & _
						"                    premi_kontanan.fddate," & _
						"                    premi_kontanan.employeecode)" & _
						"    WHERE     fcba = '"& v_txtFcba &"'"  
						

  
	if  len(trim(v_txtfcba)) > 0 then
		strGrnInv = strGrnInv & " AND fcba = '"& v_txtfcba &"'"
	end if
	if  len(trim(v_txtdatefrom)) > 0 then
		strGrnInv = strGrnInv & " AND fddate >= TO_DATE ('"& v_txtdatefrom &"', 'mm/dd/yyyy')"
		txtSelect = " From " & FormatDateTime(v_txtdatefrom,1)
	end if
	if  len(trim(v_txtdateuntil)) > 0 then
		strGrnInv = strGrnInv & " AND fddate <= TO_DATE ('"& v_txtdateuntil &"', 'mm/dd/yyyy')"
		txtSelect = txtSelect & " To " & FormatDateTime(v_txtdateuntil,1)
	end if
	if  len(trim(v_txtempcode)) > 0 then
		strGrnInv = strGrnInv & " AND employeecode ='"& v_txtempcode &"'"
		txtSelect = txtSelect & " Employee Code " & v_txtempcode
	end if
	if  len(trim(v_txtsectionname)) > 0 then
		strGrnInv = strGrnInv & " AND employee_api.get_sectionname (fcba, employeecode) ='"& v_txtsectionname &"'"
		txtSelect = txtSelect & " Division " & v_txtsectionname
	end if 
	if  len(trim(v_txtgangcode)) > 0 then
		strGrnInv = strGrnInv & " AND  employee_api.get_gangcode (fcba, employeecode) ='"& v_txtgangcode &"'"
		txtSelect = txtSelect & " Gang " & v_txtgangcode
	end if 


	' if  cdbl(to_nol(compare1)) = 1 then
	'     strGrnInv	=	strGrnInv & " and attendance_gad.gad_employeecode != vehicleactivity.vh_employeecode"
	' end if
	' if  cdbl(to_nol(compare1)) = 2 then
	'     strGrnInv	=	strGrnInv & " and attendance_gad.gad_employeecode != harvestingspb.harvesting_employeecode" 
	' end if
	' if  cdbl(to_nol(compare1)) = 3 then
	'     strGrnInv	=	strGrnInv & " and vehicleactivity.vh_employeecode != harvestingspb.harvesting_employeecode" 
	' end if

	' strGrnInv	=	strGrnInv &	" GROUP BY" & _
	' 							"     FCBA" & _
	' 							"     , NOPO" & _
	' 							"     , STOCKCODE"
								
	strGrnInv	=	strGrnInv & " ORDER BY employeecode, division, fddate" 

	' response.write strGrnInv	
	Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
	rsGrnInv.Open strGrnInv, connstr, adOpenStatic 
	
%>


<html>

<head>
<title>ePlaS Reporting - Premi Kontanan</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<link href="/plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="/plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="/plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="/plugins/responsivegridsystem/12cols.css" media="all">
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
				<table id="form-table" class="centered-margin">
					<tr>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">KEBUN</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">DARI</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">SAMPAI</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">DIVISION</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">GANG CODE</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">EMPLOYEE CODE</font>
						</td>
					</tr>
					<tr>						
						<td>
							<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
							<option></option>
							<%
							Set rsFcba = Server.CreateObject("ADODB.Recordset")
							rsFcba.CursorLocation = adUseClient
							strsql =   " select fccode fcba from businessunit " & _
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
						<td width="4">
							<input class="datepicker" type="text" maxlength="10" name="txtdatefrom" size="10" value="<%=v_txtdatefrom%>">
						</td> 
						<td width="2">
							<input class="datepicker" type="text" maxlength="10" name="txtdateuntil" size="10" value="<%=v_txtdateuntil%>">
						</td>  
						<td>
							<select name="txtsectionname" >
								<option value="">-- ALL --</option>
									<%
									Dim rssectionname, strsectionname
									strsectionname	=	"select fccode, fcname from sectionname where fccode in ('AFD-01','AFD-02','AFD-03','AFD-04','AFD-05','AFD-06','AFD-07','AFD-08','AFD-09','AFD-10','AFD-11','AFD-12','AFD-13') order by fccode"
									
									Set rssectionname = Server.CreateObject("ADODB.Recordset")
									rssectionname.Open strsectionname, connstr, adOpenStatic 
									
									if not rssectionname.eof then
									Do While Not rssectionname.eof
									if rssectionname("fccode") = v_txtsectionname then%>
											<option value="<%=rssectionname("fccode")%>" Selected><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option> 
										 <%Else%>
											<option value="<%=rssectionname("fccode")%>"><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option>  
										<%End if
										rssectionname.movenext
									loop
									end if
									rssectionname.Close
									set rssectionname = Nothing
									%>
							</select>
						</td> 
						<td>
							<select name="txtgangcode" >
								<option value="">-- ALL --</option>
									<% 
									strsectionname	=	"select fccode, fcname from gang where fcba = '"& v_txtFcba &"' order by fccode"
									
									Set rssectionname = Server.CreateObject("ADODB.Recordset")
									rssectionname.Open strsectionname, connstr, adOpenStatic 
									
									if not rssectionname.eof then
									Do While Not rssectionname.eof
									if rssectionname("fccode") = v_txtgangcode then%>
											<option value="<%=rssectionname("fccode")%>" Selected><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option> 
										 <%Else%>
											<option value="<%=rssectionname("fccode")%>"><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option>  
										<%End if
										rssectionname.movenext
									loop
									end if
									rssectionname.Close
									set rssectionname = Nothing
									%>
							</select>
						</td> 
						<td>
							<select name="txtempcode" class="chosen-select" >
								<option value="">--ALL--</option>
									<% 
									strsectionname	=	" select fccode, fcname from employee where fcba = '"& v_txtfcba &"' order by fcCode"
									
									Set rssectionname = Server.CreateObject("ADODB.Recordset")
									rssectionname.Open strsectionname, connstr, adOpenStatic 
									
									if not rssectionname.eof then
									Do While Not rssectionname.eof
									if rssectionname("fccode") = v_txtempcode then%>
											<option value="<%=rssectionname("fccode")%>" Selected><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%></option> 
										 <%Else%>
											<option value="<%=rssectionname("fccode")%>"><%=rssectionname("fccode")%> : <%=rssectionname("fcname")%> </option>  
										<%End if
										rssectionname.movenext
									loop
									end if
									rssectionname.Close
									set rssectionname = Nothing
									%>
							</select>
						</td>
						<!--
						<td>
							<font face="verdana" size="1">VEHICLE CODE</font>
						</td>
						<td width="2">
							<input type="text" name="txtvehiclecode" value="<%=txtvehiclecode%>">
						</td> -->
						<td>
							<input type="submit" value="View" />
						</td>
					</tr>
				</table>
			</div>
			</form>
			
			</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead>
		<tr>
	    <th colspan="22" style="background-color:#FFFFFF; border:none;">  
		<div id="header-wrapper">
			<div id="head_judul" style="float:left; text-align:left;">
					<h4>PT. <%=CompanyName%></h4>
					<h5><%=fcbaname%></h5>
				</div>
				
				<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
					<h2>LAPORAN PREMI KONTANAN</h2>
	                <h4>Selected By : <%=txtselect%></h4>
	                <br/>
				</div>
			</div>
	</th>
			</tr>
		<tr>
			<th style="background-color:#00A300;color:#fff;">Employee Code</th>
			<th style="background-color:#00A300;color:#fff;">Employee Name</th>
			<th style="background-color:#00A300;color:#fff;">Date</th>
			<th style="background-color:#00A300;color:#fff;">Division</th>
			<th style="background-color:#00A300;color:#fff;">Gang Code</th>
			<th style="background-color:#00A300;color:#fff;">Total Janjang</th>
			<th style="background-color:#00A300;color:#fff;">Kg TBS Extra Harian</th>
			<th style="background-color:#00A300;color:#fff;">Premi Lebih Borong TBS Harian</th>
			<th style="background-color:#00A300;color:#fff;">Total Bucket</th>
			<th style="background-color:#00A300;color:#fff;">Kg Brondolan Harian</th>
			<th style="background-color:#00A300;color:#fff;">Premi Brondolan Harian</th>
			<th style="background-color:#00A300;color:#fff;">Total Premi Kontanan</th>
		</tr>
	</thead>
<tbody>
<%if rsGrnInv.BOF and rsGrnInv.EOF then%>
		<tr>
			<td colspan="21" align="left">NO DATA AVAILABLE</td>
		</tr>
<%end if
		NO = 1
		v_total_tbs = 0
		v_premi_lebih_borong_tbs_harian = 0
		v_premi_lebih_borong_tbs_harian = 0
		v_total_bucket = 0
		v_kg_brondolan_harian = 0
		v_premi_brondolan_harian = 0
		v_total_premi_kontanan = 0 
		while not rsGrnInv.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
		<tr class="<%=cClass%>">
			<td align="center"><%=rsGrnInv("employeecode")%></td>
			<td align="left"><%=rsGrnInv("employeename")%></td>
			<td align="center"><%=rsGrnInv("fddate")%></td>
			<td align="center"><%=rsGrnInv("division")%></td>
			<td align="center"><%=rsGrnInv("gangcode")%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("total_tbs"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("kg_tbs_extra_harian"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("premi_lebih_borong_tbs_harian"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("total_bucket"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("kg_brondolan_harian"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("premi_brondolan_harian"))),2)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("total_premi_kontanan"))),2)%></td> 
		</tr>
<%			
			v_total_tbs = v_total_tbs + cdbl(to_nol(rsGrnInv("total_tbs")))
			v_kg_tbs_extra_harian = v_kg_tbs_extra_harian + cdbl(to_nol(rsGrnInv("kg_tbs_extra_harian")))	
			v_premi_lebih_borong_tbs_harian = v_premi_lebih_borong_tbs_harian + cdbl(to_nol(rsGrnInv("premi_lebih_borong_tbs_harian")))
			v_total_bucket = v_total_bucket + cdbl(to_nol(rsGrnInv("total_bucket")))
			v_kg_brondolan_harian = v_kg_brondolan_harian + cdbl(to_nol(rsGrnInv("kg_brondolan_harian")))	
			v_premi_brondolan_harian = v_premi_brondolan_harian + cdbl(to_nol(rsGrnInv("premi_brondolan_harian")))	
			v_total_premi_kontanan = v_total_premi_kontanan + cdbl(to_nol(rsGrnInv("total_premi_kontanan")))	 

			rsGrnInv.MoveNext
		Wend
	rsGrnInv.Close
	set rsGrnInv = Nothing
%>
		<tr bgcolor="#F4F4A0"> 
			<td align="center" colspan="5"><font style="font-weight: bold;">TOTAL</font></td>				
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_total_tbs)),0)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_kg_tbs_extra_harian)),0)%></font></td>  
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_premi_lebih_borong_tbs_harian)),0)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_total_bucket)),0)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_kg_brondolan_harian)),0)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_premi_brondolan_harian)),0)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_total_premi_kontanan)),0)%></font></td>  
		</tr>
	</tbody>
</table>

<!--#include file="../Bottom_Report.asp"-->
<script src="/js/jquery.inputmask.bundle.min.js"></script>
<script src="/js/inputmask.binding.js"></script>
<script src="/plugins/chosen/chosen.jquery.min.js"></script>
<script type="text/javascript">	
	// $(".chosen-select").empty();
	
	var config = {
	  '.chosen-select'           : {},
	  '.chosen-select-deselect'  : {allow_single_deselect:true},
	  '.chosen-select-no-single' : {disable_search_threshold:10},
	  '.chosen-select-no-results': {no_results_text:'Oops, nothing found!'},
	  '.chosen-select-width'     : {width:"45%"}
	}
	for (var selector in config) {
	  $(selector).chosen(config[selector]);
	}
	// $(".chosen-select").trigger("chosen:updated");
	
</script>t