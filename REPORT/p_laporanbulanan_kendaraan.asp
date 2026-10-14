<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
	
	ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

	v_txtFcba = Request("txtFcba")

	varDefaultBa = Session("fcBA")

	v_txtvehicle = Request("txtvehicle")
	v_txtYear = Request("txtYear")
	v_txtMonth = Request("txtMonth")
	v_txtown = Request("txtown")
	v_txtUOM = Request("txtUOM")
	
	If "" & v_txtYear = "" Then v_txtYear = Year(Session("v_Costbook"))
	If "" & v_txtMonth = "" Then v_txtMonth = Month(Session("v_Costbook"))
	If "" & txttdate = "" Then txttdate = textDate(date)
	
	If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
		v_txtFcba = varDefaultBa
	End If	
		
 '   if  v_txtfcba <> "" then
 '       whereCons = whereCons & "  AND fp.fcba = '" & v_txtfcba & "'"
	' end if
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

	' Filter UOM Added By SAG 21 Agu 2019
			 
	   if v_txtUOM	<> "--ALL--"  then
	      whereCons = wherecons & " AND VEHICLEGROUP_API.get_units (VEHICLE_API.get_vehiclegroupcode (vehiclecode, fcba),fcba) = '" & v_txtUOM & "' "
	   end if 
	 
	' End Filter UOM BY SAG
	
	strGrnInv		=	"   SELECT fcba," & _
						"          bulan," & _
						"          tahun," & _
						"          vehiclecode," & _
						"		   units," & _
						"          employee_api.get_employeename(fcba, driver) drivername," & _
						"          registrationno," & _
						"          driver," & _
						"          km_bi," & _
						"          liter_bi," & _
						"          ratio_bi," & _
						"          trip_bi," & _
						"          hke_bi," & _
						"          he_bi," & _
						"          downtime_bi," & _
						"          km_sbi," & _
						"          liter_sbi," & _
						"          ratio_sbi," & _
						"          trip_sbi," & _
						"          hke_sbi," & _
						"          he_sbi," & _
						"          downtime_sbi FROM (" & _
						"   SELECT fcba," & _
						"          bulan," & _
						"          tahun," & _
						"          vehiclecode," & _
						"		   VEHICLEGROUP_API.get_units (VEHICLE_API.get_vehiclegroupcode (vehiclecode, fcba),fcba) units," & _
						"          registrationno," & _
						"          MAX (driver) AS driver," & _
						"          MAX (km_bi) AS km_bi," & _
						"          MAX (liter_bi) AS liter_bi," & _
						"          MAX (ratio_bi) AS ratio_bi," & _
						"          MAX (trip_bi) AS trip_bi," & _
						"          MAX (hke_bi) AS hke_bi," & _
						"          MAX (he_bi) AS he_bi," & _
						"          MAX (downtime_bi) AS downtime_bi," & _
						"          MAX (km_sbi) AS km_sbi," & _
						"          MAX (liter_sbi) AS liter_sbi," & _
						"          MAX (ratio_sbi) AS ratio_sbi," & _
						"          MAX (trip_sbi) AS trip_sbi," & _
						"          MAX (hke_sbi) AS hke_sbi," & _
						"          MAX (he_sbi) AS he_sbi," & _
						"          MAX (downtime_sbi) AS downtime_sbi" & _
						"     FROM (SELECT fcba," & _
						"                  bulan," & _
						"                  tahun," & _
						"                  vehiclecode," & _
						"                  registrationno," & _
						"                  driver," & _
						"                  km_bi," & _
						"                  liter_bi," & _
						"                  ratio_bi," & _
						"                  trip_bi," & _
						"                  hke_bi," & _
						"                  he_bi," & _
						"                  downtime_bi," & _
						"                  0 AS km_sbi," & _
						"                  0 AS liter_sbi," & _
						"                  0 AS ratio_sbi," & _
						"                  0 AS trip_sbi," & _
						"                  0 AS hke_sbi," & _
						"                  0 AS he_sbi," & _
						"                  0 AS downtime_sbi" & _
						"             FROM (SELECT km_bi.fcba," & _
						"                          km_bi.bulan," & _
						"                          km_bi.tahun," & _
						"                          km_bi.vehiclecode," & _
						"                          km_bi.registrationno," & _
						"                          km_bi.driver," & _
						"                          km_bi.km_bi," & _
						"                          liter_bi.liter_bi," & _
						"                          CASE" & _
						"                             WHEN liter_bi.liter_bi = 0" & _
						"                             THEN 0" & _
						"                             WHEN km_bi.km_bi = 0" & _
						"                             THEN 0" & _
						"                             WHEN km_bi.vehicletypecode = 'KR'" & _
						"                             THEN" & _
						"                                ROUND ( (km_bi.km_bi / liter_bi.liter_bi), 2)" & _
						"                             ELSE" & _
						"                                ROUND ( (liter_bi.liter_bi / km_bi.km_bi), 2)" & _
						"                          END" & _
						"                             AS ratio_bi," & _
						"                          km_bi.trip_bi," & _
						"                          hke_bi.hke_bi," & _
						"                          he_bi.he_bi," & _
						"                          CASE" & _
						"                             WHEN hke_bi > he_bi.he_bi THEN 0" & _
						"                             ELSE ROUND ( (100 - ( (hke_bi / he_bi) * 100)), 2)" & _
						"                          END" & _
						"                             AS downtime_bi" & _
						"                     FROM (  SELECT vehicleactivity.fcba," & _
						"                                    TO_CHAR (vehicleactivity.vehicledate, 'mm')" & _
						"                                       AS bulan," & _
						"                                    TO_CHAR (vehicleactivity.vehicledate, 'yyyy')" & _
						"                                       AS tahun," & _
						"                                    vehicleactivity.vehiclecode," & _
						"                                    vehiclegroup.vehicletypecode," & _
						"                                    vehicle.registrationno," & _
						"                                    MAX (vehicleactivity.employeecode) AS driver," & _
						"                                    SUM (vehicleactivity.unit) AS km_bi," & _
						"                                    SUM (vehicleactivity.trip) AS trip_bi" & _
						"                               FROM vehicleactivity" & _
						"                                    INNER JOIN vehicle" & _
						"                                       ON (    vehicleactivity.fcba = vehicle.fcba" & _
						"                                           AND vehicleactivity.vehiclecode =" & _
						"                                                  vehicle.fccode)" & _
						"                                    INNER JOIN vehiclegroup" & _
						"                                       ON     (    vehicle.fcba =" & _
						"                                                      vehiclegroup.fcba" & _
						"                                               AND vehicle.vehiclegroupcode =" & _
						"                                                      vehiclegroup.vehiclegroupcode)" & _
						"                                          AND TO_NUMBER(TO_CHAR (" & _
						"                                                 vehicleactivity.vehicledate," & _
						"                                                 'mm')) = '"& v_txtMonth &"'" & _
						"                                          AND TO_NUMBER(TO_CHAR (" & _
						"                                                 vehicleactivity.vehicledate," & _
						"                                                 'yyyy')) = '"& v_txtYear &"'" & _
						"                           GROUP BY vehicleactivity.fcba," & _
						"                                    TO_CHAR (vehicleactivity.vehicledate, 'mm')," & _
						"                                    TO_CHAR (vehicleactivity.vehicledate," & _
						"                                             'yyyy')," & _
						"                                    vehicleactivity.vehiclecode," & _
						"                                    vehiclegroup.vehicletypecode," & _
						"                                    vehicle.registrationno) km_bi" & _
						"                          INNER JOIN" & _
						"                          (  SELECT issuedetails.fcba," & _
						"                                    TO_CHAR (issue.sivdate, 'mm') AS bulan," & _
						"                                    TO_CHAR (issue.sivdate, 'yyyy') AS tahun," & _
						"                                    issuedetails.locationcode," & _
						"                                    SUM (issuedetails.quantity) AS liter_bi" & _
						"                               FROM issuedetails" & _
						"                                    INNER JOIN issue" & _
						"                                       ON (    issuedetails.fcba = issue.fcba" & _
						"                                           AND issuedetails.sivcode =" & _
						"                                                  issue.sivcode)" & _
						"                              WHERE     TO_NUMBER(TO_CHAR (issue.sivdate, 'mm')) = '"& v_txtMonth &"'" & _
						"                                    AND TO_NUMBER(TO_CHAR (issue.sivdate, 'yyyy')) = '"& v_txtYear &"'" & _
						"									 AND issuedetails.stockcode = '302.B.0001.0001'" & _
						"                           GROUP BY issuedetails.fcba," & _
						"                                    TO_CHAR (issue.sivdate, 'mm')," & _
						"                                    TO_CHAR (issue.sivdate, 'yyyy')," & _
						"                                    issuedetails.locationcode) liter_bi" & _
						"                             ON (    km_bi.fcba = liter_bi.fcba" & _
						"                                 AND km_bi.vehiclecode = liter_bi.locationcode" & _
						"                                 AND km_bi.bulan = liter_bi.bulan" & _
						"                                 AND km_bi.tahun = liter_bi.tahun)" & _
						"                          INNER JOIN" & _
						"                          (  SELECT fcba," & _
						"                                    TO_CHAR (vehicledate, 'mm') AS bulan," & _
						"                                    TO_CHAR (vehicledate, 'yyyy') AS tahun," & _
						"                                    vehiclecode," & _
						"                                    SUM (hke) AS hke_bi" & _
						"                               FROM (SELECT DISTINCT vehicleactivity.fcba," & _
						"                                                     vehicleactivity.vehicledate," & _
						"                                                     vehicleactivity.vehiclecode," & _
						"                                                     1 AS hke" & _
						"                                       FROM vehicleactivity" & _
						"                                      WHERE     TO_NUMBER(TO_CHAR (" & _
						"                                                   vehicleactivity.vehicledate," & _
						"                                                   'mm')) = '"& v_txtMonth &"'" & _
						"                                            AND TO_NUMBER(TO_CHAR (" & _
						"                                                   vehicleactivity.vehicledate," & _
						"                                                   'yyyy')) = '"& v_txtYear &"')" & _
						"                           GROUP BY fcba," & _
						"                                    TO_CHAR (vehicledate, 'mm')," & _
						"                                    TO_CHAR (vehicledate, 'yyyy')," & _
						"                                    vehiclecode) hke_bi" & _
						"                             ON (    km_bi.fcba = hke_bi.fcba" & _
						"                                 AND km_bi.vehiclecode = hke_bi.vehiclecode" & _
						"                                 AND km_bi.bulan = hke_bi.bulan" & _
						"                                 AND km_bi.tahun = hke_bi.tahun)" & _
						"                          INNER JOIN" & _
						"                          (  SELECT fcba," & _
						"                                    TO_CHAR (workingdate, 'mm') AS bulan," & _
						"                                    TO_CHAR (workingdate, 'yyyy') AS tahun," & _
						"                                    COUNT (workingdate) AS he_bi" & _
						"                               FROM budget_transaction_workingday" & _
						"                              WHERE     TO_NUMBER(TO_CHAR (workingdate, 'mm')) = '"& v_txtMonth &"'" & _
						"                                    AND TO_NUMBER(TO_CHAR (workingdate, 'yyyy')) = '"& v_txtYear &"'" & _
						"                           GROUP BY fcba," & _
						"                                    TO_CHAR (workingdate, 'mm')," & _
						"                                    TO_CHAR (workingdate, 'yyyy')) he_bi" & _
						"                             ON (    km_bi.fcba = he_bi.fcba" & _
						"                                 AND km_bi.bulan = he_bi.bulan" & _
						"                                 AND km_bi.tahun = he_bi.tahun))" & _
						"           UNION ALL" & _
						"           SELECT fcba," & _
						"                  bulan," & _
						"                  tahun," & _
						"                  vehiclecode," & _
						"                  registrationno," & _
						"                  driver," & _
						"                  0 AS km_bi," & _
						"                  0 AS liter_bi," & _
						"                  0 AS ratio_bi," & _
						"                  0 AS trip_bi," & _
						"                  0 AS hke_bi," & _
						"                  0 AS he_bi," & _
						"                  0 AS downtime_bi," & _
						"                  km_sbi," & _
						"                  liter_sbi," & _
						"                  ratio_sbi," & _
						"                  trip_sbi," & _
						"                  hke_sbi," & _
						"                  he_sbi," & _
						"                  downtime_sbi" & _
						"             FROM (SELECT km_sbi.fcba," & _
						"                          km_sbi.bulan," & _
						"                          km_sbi.tahun," & _
						"                          km_sbi.vehiclecode," & _
						"                          km_sbi.registrationno," & _
						"                          km_sbi.driver," & _
						"                          km_sbi.km_sbi," & _
						"                          liter_sbi.liter_sbi," & _
						"                          CASE" & _
						"                             WHEN liter_sbi.liter_sbi = 0" & _
						"                             THEN 0" & _
						"                             WHEN km_sbi.km_sbi = 0" & _
						"                             THEN 0" & _
						"                             WHEN km_sbi.vehicletypecode = 'KR'" & _
						"                             THEN" & _
						"                                ROUND ( (km_sbi.km_sbi / liter_sbi.liter_sbi)," & _
						"                                       2)" & _
						"                             ELSE" & _
						"                                ROUND ( (liter_sbi.liter_sbi / km_sbi.km_sbi)," & _
						"                                       2)" & _
						"                          END" & _
						"                             AS ratio_sbi," & _
						"                          km_sbi.trip_sbi," & _
						"                          hke_sbi.hke_sbi," & _
						"                          he_sbi.he_sbi," & _
						"                          CASE" & _
						"                             WHEN hke_sbi > he_sbi" & _
						"                             THEN 0" & _
						"                             ELSE" & _
						"                                ROUND ( (100 - ( (hke_sbi / he_sbi) * 100)), 2)" & _
						"                          END" & _
						"                             AS downtime_sbi" & _
						"                     FROM (  SELECT vehicleactivity.fcba," & _
						"                                    MAX (" & _
						"                                       TO_CHAR (vehicleactivity.vehicledate, 'mm'))" & _
						"                                       AS bulan," & _
						"                                    MAX (" & _
						"                                       TO_CHAR (vehicleactivity.vehicledate," & _
						"                                                'yyyy'))" & _
						"                                       AS tahun," & _
						"                                    vehicleactivity.vehiclecode," & _
						"                                    vehiclegroup.vehicletypecode," & _
						"                                    vehicle.registrationno," & _
						"                                    MAX (vehicleactivity.employeecode) AS driver," & _
						"                                    SUM (vehicleactivity.unit) AS km_sbi," & _
						"                                    SUM (vehicleactivity.trip) AS trip_sbi" & _
						"                               FROM vehicleactivity" & _
						"                                    INNER JOIN vehicle" & _
						"                                       ON (    vehicleactivity.fcba = vehicle.fcba" & _
						"                                           AND vehicleactivity.vehiclecode =" & _
						"                                                  vehicle.fccode)" & _
						"                                    INNER JOIN vehiclegroup" & _
						"                                       ON     (    vehicle.fcba =" & _
						"                                                      vehiclegroup.fcba" & _
						"                                               AND vehicle.vehiclegroupcode =" & _
						"                                                      vehiclegroup.vehiclegroupcode)" & _
						"                                          AND TO_NUMBER(TO_CHAR (" & _
						"                                                  vehicleactivity.vehicledate," & _
						"                                                  'MM')) <= '"& v_txtMonth &"'" & _
						"                                          AND TO_NUMBER(TO_CHAR (" & _
						"                                                  vehicleactivity.vehicledate," & _
						"                                                  'YYYY')) = '"& v_txtYear &"'" & _
						"                           GROUP BY vehicleactivity.fcba," & _
						"                                    vehicleactivity.vehiclecode," & _
						"                                    vehiclegroup.vehicletypecode," & _
						"                                    vehicle.registrationno) km_sbi" & _
						"                          INNER JOIN" & _
						"                          (  SELECT issuedetails.fcba," & _
						"                                    MAX (TO_CHAR (issue.sivdate, 'mm')) AS bulan," & _
						"                                    MAX (TO_CHAR (issue.sivdate, 'yyyy')) AS tahun," & _
						"                                    issuedetails.locationcode," & _
						"                                    SUM (issuedetails.quantity) AS liter_sbi" & _
						"                               FROM issuedetails" & _
						"                                    INNER JOIN issue" & _
						"                                       ON (    issuedetails.fcba = issue.fcba" & _
						"                                           AND issuedetails.sivcode =" & _
						"                                                  issue.sivcode)" & _
						"                              WHERE     TO_NUMBER(TO_CHAR (issue.sivdate, 'MM')) <= '"& v_txtMonth &"'" & _
						"                                    AND TO_NUMBER(TO_CHAR (issue.sivdate, 'YYYY')) = '"& v_txtYear &"'" & _
						"									 AND issuedetails.stockcode = '302.B.0001.0001'" & _
						"                           GROUP BY issuedetails.fcba, issuedetails.locationcode)" & _
						"                          liter_sbi" & _
						"                             ON (    km_sbi.fcba = liter_sbi.fcba" & _
						"                                 AND km_sbi.vehiclecode = liter_sbi.locationcode" & _
						"                                 AND km_sbi.bulan = liter_sbi.bulan" & _
						"                                 AND km_sbi.tahun = liter_sbi.tahun)" & _
						"                          INNER JOIN" & _
						"                          (  SELECT fcba," & _
						"                                    MAX (TO_CHAR (vehicledate, 'mm')) AS bulan," & _
						"                                    MAX (TO_CHAR (vehicledate, 'yyyy')) AS tahun," & _
						"                                    vehiclecode," & _
						"                                    SUM (hke) AS hke_sbi" & _
						"                               FROM (SELECT DISTINCT vehicleactivity.fcba," & _
						"                                                     vehicleactivity.vehicledate," & _
						"                                                     vehicleactivity.vehiclecode," & _
						"                                                     1 AS hke" & _
						"                                       FROM vehicleactivity" & _
						"                                      WHERE     TO_NUMBER(TO_CHAR (" & _
						"                                                   vehicleactivity.vehicledate," & _
						"                                                   'MM')) <= '"& v_txtMonth &"'" & _
						"                                            AND TO_NUMBER(TO_CHAR (" & _
						"                                                   vehicleactivity.vehicledate," & _
						"                                                   'YYYY')) = '"& v_txtYear &"')" & _
						"                           GROUP BY fcba, vehiclecode) hke_sbi" & _
						"                             ON (    km_sbi.fcba = hke_sbi.fcba" & _
						"                                 AND km_sbi.vehiclecode = hke_sbi.vehiclecode" & _
						"                                 AND km_sbi.bulan = hke_sbi.bulan" & _
						"                                 AND km_sbi.tahun = hke_sbi.tahun)" & _
						"                          INNER JOIN" & _
						"                          (  SELECT fcba," & _
						"                                    MAX (TO_CHAR (workingdate, 'mm')) AS bulan," & _
						"                                    MAX (TO_CHAR (workingdate, 'yyyy')) AS tahun," & _
						"                                    COUNT (workingdate) AS he_sbi" & _
						"                               FROM budget_transaction_workingday" & _
						"                              WHERE     TO_NUMBER(TO_CHAR (workingdate, 'MM')) <= '"& v_txtMonth &"'" & _
						"                                    AND TO_NUMBER(TO_CHAR (workingdate, 'YYYY')) = '"& v_txtYear &"'" & _
						"                           GROUP BY fcba) he_sbi" & _
						"                             ON (    km_sbi.fcba = he_sbi.fcba" & _
						"                                 AND km_sbi.bulan = he_sbi.bulan" & _
						"                                 AND km_sbi.tahun = he_sbi.tahun)))" & _
						"							WHERE fcba ='"& v_txtFcba &"'" & whereCons & ""
						
						

	if  len(trim(v_txtvehicle)) > 0 then
		strGrnInv = strGrnInv & " AND vehiclecode ='"& v_txtvehicle &"'"
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

	strGrnInv	=	strGrnInv &	" GROUP BY fcba," & _
								"          bulan," & _
								"          tahun," & _ 
								"          vehiclecode," & _
								"          registrationno"  
								
	strGrnInv	=	strGrnInv &	" ORDER BY fcba," & _
								"          bulan," & _
								"          tahun," & _
								"          vehiclecode)" 

	' response.write strGrnInv	
	Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
	rsGrnInv.Open strGrnInv, connstr, adOpenStatic 
	
%>


<html>

<head>
<title>ePlaS Reporting - Laporan Bulanan Kendaraan</title>
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
							<font face="verdana" size="1" style="font-weight: bold;">BUSINESS UNIT</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">YEAR</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">MONTH</font>
						</td> 
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">VEHICLE CODE</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">UOM</font>
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
						<td width="2">
							<select name="txtMonth">
							<%For i = 1 to 12%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
							<%Next%>
							</select>
						</td> 
						<td>
							<select name="txtvehicle" class="chosen-select" >
								<option value="">--ALL--</option>
									<% 
									strsectionname	=	" select fccode, fcname from vehicle where fcba = '"& v_txtfcba &"' order by fcCode"
									
									Set rssectionname = Server.CreateObject("ADODB.Recordset")
									rssectionname.Open strsectionname, connstr, adOpenStatic 
									
									if not rssectionname.eof then
									Do While Not rssectionname.eof
									if rssectionname("fccode") = v_txtvehicle then%>
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
						<td align="left">
							<select name="txtUOM" >
								<option value="--ALL--">--ALL--</option>
								<option value="KM" <% If cstr(""&v_txtUOM) = "KM" Then %> Selected <%End if%>>KM</option> 
								<option value="HM" <% If cstr(""&v_txtUOM) = "HM" Then %> Selected <%End if%>>HM</option>  
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
<thead>
    <tr>
        <th colspan="16" style="background-color:#FFFFFF; border:none;">
            <div id="header-wrapper">
                <div id="head_judul" style="float:left; text-align:left;">
                    <h4>PT. <%=CompanyName%></h4>
                    <h5><%=fcbaname%></h5>
                </div>

                <div id="subheader" class="centered" tyle="width:50%; margin:auto;">
                    <h2>LAPORAN BULANAN KENDARAAN</h2>
                    <h4>PERIODE: Year <%=v_txtyear%> Month <%=monthname(v_txtMonth)%> </h4>
                    <br/>
                </div>
            </div>
        </th>
    </tr>
    <tr> 
        <th style="background-color:#00A300;color:#fff;" rowspan="2">KODE KENDARAAN</th>
        <th style="background-color:#00A300;color:#fff;" rowspan="2">KETERANGAN</th> 
    	<th style="background-color:#00A300;color:#fff;" colspan="7">BI</th>
        <th style="background-color:#00A300;color:#fff;" colspan="7">SBI</th>
    </tr>
    <tr> 
    	<% If v_txtUOM = "KM" Then%>
			<th style="background-color:#00A300;color:#fff;"  >KM</th>
		<% Else If v_txtUOM = "HM" Then %>
			<th style="background-color:#00A300;color:#fff;"  >HM</th>
		<% Else %>
			<th style="background-color:#00A300;color:#fff;"  >KM/HM</th>
		<% End If %>
		<% End If %>
        <th style="background-color:#00A300;color:#fff;" >LTR</th>
        <th style="background-color:#00A300;color:#fff;" >RATIO</th>
        <th style="background-color:#00A300;color:#fff;" >TRIP</th>
        <th style="background-color:#00A300;color:#fff;" >HKE</th>
        <th style="background-color:#00A300;color:#fff;" >HE</th>
        <th style="background-color:#00A300;color:#fff;" >DOWN TIME (%)</th>

        <% If v_txtUOM = "KM" Then%>
			<th style="background-color:#00A300;color:#fff;"  >KM</th>
		<% Else If v_txtUOM = "HM" Then %>
			<th style="background-color:#00A300;color:#fff;"  >HM</th>
		<% Else %>
			<th style="background-color:#00A300;color:#fff;"  >KM/HM</th>
		<% End If %>
		<% End If %>
        <th style="background-color:#00A300;color:#fff;" >LTR</th>
        <th style="background-color:#00A300;color:#fff;" >RATIO</th>
        <th style="background-color:#00A300;color:#fff;" >TRIP</th>
        <th style="background-color:#00A300;color:#fff;" >HKE</th>
        <th style="background-color:#00A300;color:#fff;" >HE</th>
        <th style="background-color:#00A300;color:#fff;" >DOWN TIME (%)</th>
    </tr>
</thead>
<tbody>
<%if rsGrnInv.BOF and rsGrnInv.EOF then%>
		<tr>
			<td colspan="16" align="left">NO DATA AVAILABLE</td>
		</tr>
<%end if
		NO = 1
		v_premi = 0
		 
		while not rsGrnInv.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
		<tr class="<%=cClass%>"> 
			<td align="center"><%=rsGrnInv("vehiclecode")%></td>
			<td align="center"><%=rsGrnInv("registrationno")%></td> 

			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("km_bi"))),0)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("liter_bi"))),0)%></td>
			<td align="center"><%=formatnumber(cdbl(to_nol(rsGrnInv("ratio_bi"))),2)%></td>
			<td align="center"><%=formatnumber(cdbl(to_nol(rsGrnInv("trip_bi"))),0)%></td>
			<td align="center"><%=formatnumber(cdbl(to_nol(rsGrnInv("hke_bi"))),0)%></td>
			<td align="center"><%=formatnumber(cdbl(to_nol(rsGrnInv("he_bi"))),0)%></td>
			<td align="center"><%=formatnumber(cdbl(to_nol(rsGrnInv("downtime_bi"))),0)%></td>

			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("km_sbi"))),0)%></td>
			<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("liter_sbi"))),0)%></td>
			<td align="center"><%=formatnumber(cdbl(to_nol(rsGrnInv("ratio_sbi"))),2)%></td>
			<td align="center"><%=formatnumber(cdbl(to_nol(rsGrnInv("trip_sbi"))),0)%></td>
			<td align="center"><%=formatnumber(cdbl(to_nol(rsGrnInv("hke_sbi"))),0)%></td>
			<td align="center"><%=formatnumber(cdbl(to_nol(rsGrnInv("he_sbi"))),0)%></td>
			<td align="center"><%=formatnumber(cdbl(to_nol(rsGrnInv("downtime_sbi"))),0)%></td>
		</tr>
<%			
			' v_premi = v_premi + cdbl(to_nol(rsGrnInv("premi"))) 

			rsGrnInv.MoveNext
		Wend
	rsGrnInv.Close
	set rsGrnInv = Nothing
%> 
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
	
</script>