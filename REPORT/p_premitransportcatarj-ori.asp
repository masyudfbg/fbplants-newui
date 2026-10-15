<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
	Dim v_txtYear, v_txtMonth
	Dim ScriptName
	Dim rsGrnInv, strGrnInv
	Dim v_txtPOCode
	Dim NO, cClass
	
	ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
	Session("CurP_Vehicle") = cmdFindAccountPeriod("Vehicle")
	Session("ClsP_Vehicle") = cmdAccountPeriodClose("Vehicle")
	Session("v_Vehicle") = cmdFindPeriod("Vehicle")
	
	If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))
	If "" & txtMonth = "" Then txtMonth = Month(Session("v_Costbook"))
	If "" & txttdate = "" Then txttdate = textDate(date)
	varDefaultBa = Session("fcBA")
	v_txtFcba = Request("txtFcba")
	
	txttdate = txtMonth & "" & txtYear
	v_txtPOCode		=	trim(request("txtPOCode"))
	compare1 = trim(request("compare1"))

	varDefaultBa = Session("fcBA")

	v_txtsectionname = request("txtsectionname")
	v_txtvehiclecode = request("txtvehiclecode")
	v_txtdatefrom = Request("txtdatefrom")
	v_txtdateuntil = Request("txtdateuntil")
	v_txtown = Request("txtown")
	v_txtempcode = Request("txtempcode")
	If (v_txtdatefrom = "" Or Len(Trim(v_txtdatefrom)) = 0) Then v_txtdatefrom = textdate(date)
	If (v_txtdateuntil = "" Or Len(Trim(v_txtdateuntil)) = 0) Then v_txtdateuntil = textdate(date)

	If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
		v_txtFcba = varDefaultBa
	End If	
	
	If len(trim(v_txtempcode)) > 0 Then 
		WhereCons = Wherecons & " AND vehicleactivity.employeecode = '"& v_txtempcode &"'"
	End If
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
	  
	strGrnInv		=	" SELECT vhact_fix.fcba," & _
						"        vhact_fix.vehicledate," & _
						"        vhact_fix.employeecode," & _
						"        emp.fcname AS employeename," & _
						"        vhact_fix.vehiclecode," & _
						"        vhact_fix.categoryname," & _
						"        vhact_fix.hasilkerja," & _
						"        vhact_fix.basis," & _
						"        vhact_fix.lebihbasis," & _
						"        vhact_fix.rate," & _
						"        CASE" & _
						"           WHEN (    cat_atas.category_true_teratas = vhact_fix.categorycode" & _
						"                 AND vhact_fix.statusbasis = 'TRUE'" & _
						"                 AND vhact_fix.daystatus = 'Hari Kerja')" & _
						"           THEN" & _
						"              lebihbasis * rate" & _
						"           ELSE" & _
						"              vhact_fix.hasilkerja * rate" & _
						"        END" & _
						"           AS premi_operator" & _
						"   FROM (SELECT fcba," & _
						"                employeecode," & _
						"                CASE" & _
						"                   WHEN harilibur IS NULL THEN 'Hari Kerja'" & _
						"                   ELSE 'Hari Libur'" & _
						"                END" & _
						"                   AS daystatus," & _
						"                vehiclegroupcode," & _
						"                vehicledate," & _
						"                vehiclecode," & _
						"                categorycode," & _
						"                categoryname," & _
						"                method_by," & _
						"                hasilkerja," & _
						"                basis," & _
						"                CASE" & _
						"                   WHEN hasilkerja - basis < 0 THEN 0" & _
						"                   ELSE hasilkerja - basis" & _
						"                END" & _
						"                   AS lebihbasis," & _
						"                CASE WHEN hasilkerja - basis < 0 THEN 'FALSE' ELSE 'TRUE' END" & _
						"                   AS statusbasis," & _
						"                rate" & _
						"           FROM (  SELECT vehicleactivity.fcba," & _
						"                          vehicleactivity.employeecode," & _
						"                          harilibur.harilibur," & _
						"                          vehiclegroup.vehiclegroupcode," & _
						"                          vehicleactivity.vehicledate," & _
						"                          vehicleactivity.vehiclecode," & _
						"                          to_number(vehicleactivitycategory.fccode) AS categorycode," & _
						"                          vehicleactivitycategory.fcname AS categoryname," & _
						"                          NVL (SUM (vehicleactivity.trip), 0) AS trip," & _
						"                          NVL (SUM (vehicleactivity.unit), 0) AS KMHM," & _
						"                          NVL (SUM (vehicleactivity.output), 0) AS output," & _
						"                          operator_category_rate.method_by," & _
						"                          CASE" & _
						"                             WHEN operator_category_rate.method_by = 'OUTPUT'" & _
						"                             THEN" & _
						"                                NVL (SUM (vehicleactivity.output), 0)" & _
						"                             WHEN operator_category_rate.method_by = 'HM'" & _
						"                             THEN" & _
						"                                NVL (SUM (vehicleactivity.unit), 0)" & _
						"                             ELSE" & _
						"                                NVL (SUM (vehicleactivity.trip), 0)" & _
						"                          END" & _
						"                             AS hasilkerja," & _
						"                          operator_category_rate.basis," & _
						"                          operator_category_rate.rate" & _
						"                     FROM vehicleactivity" & _
						"                          INNER JOIN vehicleactivitycategory" & _
						"                             ON (vehicleactivity.categorytransport =" & _
						"                                    vehicleactivitycategory.fccode)" & _
						"                          INNER JOIN" & _
						"                          (SELECT fcba," & _
						"                                  fccode," & _
						"                                  fcname," & _
						"                                  vehiclegroupcode" & _
						"                             FROM vehicle) vehicle" & _
						"                             ON (    vehicleactivity.fcba = vehicle.fcba" & _
						"                                 AND vehicleactivity.vehiclecode =" & _
						"                                        vehicle.fccode)" & _
						"                          INNER JOIN" & _
						"                          (SELECT fcba, vehiclegroupcode, description" & _
						"                             FROM vehiclegroup) vehiclegroup" & _
						"                             ON (    vehicleactivity.fcba = vehiclegroup.fcba" & _
						"                                 AND vehicle.vehiclegroupcode =" & _
						"                                        vehiclegroup.vehiclegroupcode)" & _
						"                          INNER JOIN" & _
						"                          (SELECT DISTINCT fcba," & _
						"                                           fddate," & _
						"                                           employeecode," & _
						"                                           locationcode," & _
						"                                           attendance" & _
						"                             FROM attendance_gad" & _
						"                            WHERE     fcba = '" & v_txtFcba & "'" & _
						"                                  AND locationtype = 'VH'" & _
						"                                  AND fddate >= to_date('"& v_txtdatefrom &"', 'mm/dd/yyyy')" & _
						"                                  AND fddate <= to_date('"& v_txtdateuntil &"', 'mm/dd/yyyy')" & _
						"                                  AND attendance IN ('KJ'," & _
						"                                                     'WH'," & _
						"                                                     'WS'," & _
						"                                                     'KK'," & _
						"                                                     'DL')) attendance_gad" & _
						"                             ON (    attendance_gad.fcba = vehicleactivity.fcba" & _
						"                                 AND attendance_gad.fddate =" & _
						"                                        vehicleactivity.vehicledate" & _
						"                                 AND attendance_gad.employeecode =" & _
						"                                        vehicleactivity.employeecode" & _
						"                                 AND attendance_gad.locationcode =" & _
						"                                        vehicleactivity.vehiclecode)" & _
						"                          INNER JOIN operator_category_rate" & _
						"                             ON (    vehicleactivity.fcba =" & _
						"                                        operator_category_rate.fcba" & _
						"                                 AND vehicleactivity.categorytransport =" & _
						"                                        operator_category_rate.categorytransport" & _
						"                                 AND vehicle.vehiclegroupcode =" & _
						"                                        operator_category_rate.vehicle_group" & _
						"                                 AND attendance_gad.attendance =" & _
						"                                        operator_category_rate.attendance)" & _
						"                          LEFT JOIN" & _
						"                          (SELECT DISTINCT HariLibur" & _
						"                             FROM (SELECT dt AS HariLibur" & _
						"                                     FROM (WITH dts" & _
						"                                                AS (    SELECT   DATE '"& Year(Session("v_Vehicle")) &"-01-01'" & _
						"                                                               + ROWNUM -1" & _
						"                                                                  dt" & _
						"                                                          FROM DUAL" & _
						"                                                    CONNECT BY LEVEL <= 366)" & _
						"                                           SELECT *" & _
						"                                             FROM dts" & _
						"                                            WHERE TO_CHAR (dt, 'fmday') =" & _
						"                                                     'sunday')" & _
						"                                   UNION ALL" & _
						"                                   SELECT DISTINCT holidaydate" & _
						"                                     FROM budget_transaction_holiday" & _
						"                                    WHERE fcba IN ('" & v_txtFcba & "'," & _
						"                                                   'PSE'," & _
						"                                                   'SKE'," & _
						"                                                   'ARM'," & _
						"                                                   'ARJ'))) harilibur" & _
						"                             ON (vehicleactivity.vehicledate =" & _
						"                                    harilibur.HariLibur)" & _
						"                    WHERE     vehicleactivity.fcba = '" & v_txtFcba & "'" & _
						"                          AND vehicleactivity.vehicledate >= to_date('"& v_txtdatefrom &"','mm/dd/yyyy')" & _
						"                          AND vehicleactivity.vehicledate <= to_date('"& v_txtdateuntil &"','mm/dd/yyyy')" & _
						"                          "& Wherecons &"" & _
						"                 GROUP BY vehicleactivity.fcba," & _
						"                          vehicleactivity.employeecode," & _
						"                          attendance_gad.attendance," & _
						"                          harilibur.harilibur," & _
						"                          vehiclegroup.vehiclegroupcode," & _
						"                          vehicleactivity.vehicledate," & _
						"                          vehicleactivity.vehiclecode," & _
						"                          vehicleactivitycategory.fccode," & _
						"                          vehicleactivitycategory.fcname," & _
						"                          operator_category_rate.method_by," & _
						"                          operator_category_rate.basis," & _
						"                          operator_category_rate.rate)) vhact_fix" & _
						"        INNER JOIN" & _
						"        (SELECT fcba," & _
						"                employeecode," & _
						"                vehiclegroupcode," & _
						"                vehicledate," & _
						"                vehiclecode," & _
						"                categorycode," & _
						"                statusbasis," & _
						"                category_true_teratas" & _
						"           FROM (SELECT fcba," & _
						"                        employeecode," & _
						"                        vehiclegroupcode," & _
						"                        vehicledate," & _
						"                        vehiclecode," & _
						"                        categorycode," & _
						"                        statusbasis," & _
						"                        MIN (" & _
						"                           categorycode)" & _
						"                        OVER (" & _
						"                           PARTITION BY fcba || employeecode || vehicledate)" & _
						"                           AS category_true_teratas" & _
						"                   FROM (SELECT fcba," & _
						"                                employeecode," & _
						"                                vehiclegroupcode," & _
						"                                vehicledate," & _
						"                                vehiclecode," & _
						"                                categorycode," & _
						"                                CASE" & _
						"                                 WHEN (    daystatus = 'Hari Kerja'" & _
						"										AND hasilkerja - basis < 0)" & _
						"								  THEN" & _
						"									 'FALSE'" & _
						"								  WHEN (    daystatus = 'Hari Libur'" & _
						"										AND hasilkerja - basis < 0)" & _
						"								  THEN" & _
						"									 'TRUE'" & _
						"								  ELSE" & _
						"									 'TRUE'" & _
						"                                END" & _
						"                                   AS statusbasis" & _
						"                           FROM (  SELECT vehicleactivity.fcba," & _
						"                                          vehicleactivity.employeecode," & _
						"                                          vehiclegroup.vehiclegroupcode," & _
						"                                          vehicleactivity.vehicledate," & _
						"										   CASE " & _
						"											WHEN harilibur.harilibur IS NULL" & _
						"											THEN" & _
						"											   'Hari Kerja'" & _
						"											ELSE" & _
						"											   'Hari Libur'" & _
						"										 END" & _
						"											AS daystatus," & _
						"                                          vehicleactivity.vehiclecode," & _
						"                                          to_number(vehicleactivitycategory.fccode) AS categorycode," & _
						"                                          NVL (SUM (vehicleactivity.trip), 0)" & _
						"                                             AS trip," & _
						"                                          NVL (SUM (vehicleactivity.unit), 0)" & _
						"                                             AS KMHM," & _
						"                                          NVL (SUM (vehicleactivity.output), 0)" & _
						"                                             AS output," & _
						"                                          operator_category_rate.method_by," & _
						"                                          CASE" & _
						"                                             WHEN operator_category_rate.method_by =" & _
						"                                                     'OUTPUT'" & _
						"                                             THEN" & _
						"                                                NVL (" & _
						"                                                   SUM (vehicleactivity.output)," & _
						"                                                   0)" & _
						"                                             WHEN operator_category_rate.method_by =" & _
						"                                                     'HM'" & _
						"                                             THEN" & _
						"                                                NVL (SUM (vehicleactivity.unit)," & _
						"                                                     0)" & _
						"                                             ELSE" & _
						"                                                NVL (SUM (vehicleactivity.trip)," & _
						"                                                     0)" & _
						"                                          END" & _
						"                                             AS hasilkerja," & _
						"                                          operator_category_rate.basis" & _
						"                                     FROM vehicleactivity" & _
						"                                          INNER JOIN vehicleactivitycategory" & _
						"                                             ON (vehicleactivity.categorytransport =" & _
						"                                                    vehicleactivitycategory.fccode)" & _
						"                                          INNER JOIN" & _
						"                                          (SELECT fcba," & _
						"                                                  fccode," & _
						"                                                  fcname," & _
						"                                                  vehiclegroupcode" & _
						"                                             FROM vehicle) vehicle" & _
						"                                             ON (    vehicleactivity.fcba =" & _
						"                                                        vehicle.fcba" & _
						"                                                 AND vehicleactivity.vehiclecode =" & _
						"                                                        vehicle.fccode)" & _
						"                                          INNER JOIN" & _
						"                                          (SELECT fcba," & _
						"                                                  vehiclegroupcode," & _
						"                                                  description" & _
						"                                             FROM vehiclegroup) vehiclegroup" & _
						"                                             ON (    vehicleactivity.fcba =" & _
						"                                                        vehiclegroup.fcba" & _
						"                                                 AND vehicle.vehiclegroupcode =" & _
						"                                                        vehiclegroup.vehiclegroupcode)" & _
						"                                          INNER JOIN" & _
						"                                          (SELECT DISTINCT fcba," & _
						"                                                           fddate," & _
						"                                                           employeecode," & _
						"                                                           locationcode," & _
						"                                                           attendance" & _
						"                                             FROM attendance_gad" & _
						"                                            WHERE     fcba = '" & v_txtFcba & "'" & _
						"                                                  AND locationtype = 'VH'" & _
						"                                  				   AND fddate >= to_date('"& v_txtdatefrom &"', 'mm/dd/yyyy')" & _
						"                                  				   AND fddate <= to_date('"& v_txtdateuntil &"', 'mm/dd/yyyy')" & _
						"                                                  AND attendance IN ('KJ'," & _
						"                                                                     'WH'," & _
						"                                                                     'WS'," & _
						"                                                                     'KK'," & _
						"                                                                     'DL'))" & _
						"                                          attendance_gad" & _
						"                                             ON (    attendance_gad.fcba =" & _
						"                                                        vehicleactivity.fcba" & _
						"                                                 AND attendance_gad.fddate =" & _
						"                                                        vehicleactivity.vehicledate" & _
						"                                                 AND attendance_gad.employeecode =" & _
						"                                                        vehicleactivity.employeecode" & _
						"                                                 AND attendance_gad.locationcode =" & _
						"                                                        vehicleactivity.vehiclecode)" & _
						"                                          INNER JOIN operator_category_rate" & _
						"                                             ON (    vehicleactivity.fcba =" & _
						"                                                        operator_category_rate.fcba" & _
						"                                                 AND vehicleactivity.categorytransport =" & _
						"                                                        operator_category_rate.categorytransport" & _
						"                                                 AND vehicle.vehiclegroupcode =" & _
						"                                                        operator_category_rate.vehicle_group" & _
						"                                                 AND attendance_gad.attendance =" & _
						"                                                        operator_category_rate.attendance)" & _
						"                                         LEFT JOIN" & _
                        "                 (SELECT DISTINCT HariLibur" & _
                        "                    FROM (SELECT dt AS HariLibur" & _
                        "                            FROM (WITH dts" & _
                        "                                       AS (    SELECT   DATE '"& Year(Session("v_Vehicle")) &"-01-01'" & _
                        "                                                      + ROWNUM" & _
                        "                                                      - 1" & _
                        "                                                         dt" & _
                        "                                                 FROM DUAL" & _
                        "                                           CONNECT BY LEVEL <=" & _
                        "                                                         366)" & _
                        "                                  SELECT *" & _
                        "                                    FROM dts" & _
                        "                                   WHERE TO_CHAR (" & _
                        "                                            dt," & _
                        "                                            'fmday') =" & _
                        "                                            'sunday')" & _
                        "                          UNION ALL" & _
                        "                          SELECT DISTINCT holidaydate" & _
                        "                            FROM budget_transaction_holiday" & _
                        "                           WHERE fcba IN ('GWE'," & _
                        "                                          'PSE'," & _
                        "                                          'SKE'," & _
                        "                                          'ARM'," & _
                        "                                          'ARJ')))" & _
                        "                 harilibur" & _
                        "                    ON (vehicleactivity.vehicledate = harilibur.HariLibur)" & _
						"                    				WHERE     vehicleactivity.fcba = '" & v_txtFcba & "'" & _
						"                          			AND vehicleactivity.vehicledate >= to_date('"& v_txtdatefrom &"','mm/dd/yyyy')" & _
						"                          			AND vehicleactivity.vehicledate <= to_date('"& v_txtdateuntil &"','mm/dd/yyyy')" & _
						"                          			"& Wherecons &"" & _
						"                                 GROUP BY vehicleactivity.fcba," & _
						"                                          vehicleactivity.employeecode," & _
						"                                          vehiclegroup.vehiclegroupcode," & _
						"                                          vehicleactivity.vehicledate," & _
						"                                          vehicleactivity.vehiclecode," & _
						"                                          harilibur.harilibur," & _
						"                                          vehicleactivitycategory.fccode," & _
						"                                          operator_category_rate.method_by," & _
						"                                          operator_category_rate.basis," & _
						"                                          operator_category_rate.rate))" & _
						"                  WHERE statusbasis = 'TRUE')" & _
						"          WHERE category_true_teratas = categorycode) cat_atas" & _
						"           ON (    vhact_fix.fcba = cat_atas.fcba" & _
						"               AND vhact_fix.employeecode = cat_atas.employeecode" & _
						"               AND vhact_fix.vehicledate = cat_atas.vehicledate)" & _
						"        INNER JOIN (SELECT fcba, fccode, fcname" & _
						"                      FROM employee" & _
						"                     WHERE fcba = '" & v_txtFcba & "') emp" & _
						"           ON (    vhact_fix.fcba = emp.fcba" & _
						"               AND vhact_fix.employeecode = emp.fccode)"  & _
						" ORDER BY vhact_fix.employeecode, vhact_fix.vehicledate"

   
	' if  len(trim(v_txtempcode)) > 0 then
		' strGrnInv = strGrnInv & " AND employeecode ='"& v_txtempcode &"'"
	' end if 
								
	' strGrnInv	=	strGrnInv & "  ORDER BY fcba, fddate, employeecode)" 

	' response.write strGrnInv	
	Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
	rsGrnInv.Open strGrnInv, connstr, adOpenStatic 
	
%>


<html>

<head>
<title>ePlaS Reporting - Premi Transport Category</title>
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
							<font face="verdana" size="1" style="font-weight: bold;">EMPLOYEE CODE</font>
						</td>
					</tr>
					<tr>						
						<td>
							<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
							<%
							Dim rsFcba
							Set rsFcba = Server.CreateObject("ADODB.Recordset")
							rsFcba.CursorLocation = adUseClient
							strsql =   	" select fccode fcba from businessunit " & _
										" where " & _
										" USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL and fctype= 'E' "
							' response.write strsql			
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
	    	<th colspan="23" style="background-color:#FFFFFF; border:none;">  
				<div id="header-wrapper">
					<div id="head_judul" style="float:left; text-align:left;">
						<h4>PT. <%=CompanyName%></h4>
						<h5><%=fcbaname%></h5>
					</div>
					<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
						<h2>PREMI TRANSPORT CATEGORY</h2>
		                <br/>
					</div>
				</div>
			</th>
		</tr>
		<tr> 
			<th style="background-color:#00A300;color:#fff;">Date</th>
			<th style="background-color:#00A300;color:#fff;">Employee Code</th>
			<th style="background-color:#00A300;color:#fff;">Employee Name</th>
			<th style="background-color:#00A300;color:#fff;">Vehicle Code</th>
			<th style="background-color:#00A300;color:#fff;">Category Name</th>
			<th style="background-color:#00A300;color:#fff;">Hasil Kerja</th>
			<th style="background-color:#00A300;color:#fff;">Basis</th>
			<th style="background-color:#00A300;color:#fff;">Lebih Basis</th>
			<th style="background-color:#00A300;color:#fff;">Rate</th>
			<th style="background-color:#00A300;color:#fff;">Premi Operator</th>
		</tr>
	</thead>
	<tbody>
	<%if rsGrnInv.BOF and rsGrnInv.EOF then%>
			<tr>
				<td colspan="23" align="left">NO DATA AVAILABLE</td>
			</tr>
	<%end if
			NO = 1
			v_hasilkerja = 0
			v_lebih_basis = 0
			v_premi_operator = 0
			while not rsGrnInv.eof
				if no  mod 2 = 1 then
					cClass = "odd"
				else
					cClass = "default"
				end if
	%>
			<tr class="<%=cClass%>">
				<td align="center"><%=rsGrnInv("vehicledate")%></td>
				<td align="center"><%=rsGrnInv("employeecode")%></td>
				<td align="center"><%=rsGrnInv("employeename")%></td>
				<td align="center"><%=rsGrnInv("vehiclecode")%></td>
				<td align="left"><%=rsGrnInv("categoryname")%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("hasilkerja"))),2)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("basis"))),2)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("lebihbasis"))),2)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("rate"))),2)%></td>
				<td align="right"><%=formatnumber(cdbl(to_nol(rsGrnInv("premi_operator"))),2)%></td>
			</tr>
	<%			v_hasilkerja = v_hasilkerja + cdbl(to_nol(rsGrnInv("hasilkerja")))
				v_lebih_basis = v_lebih_basis + cdbl(to_nol(rsGrnInv("lebihbasis")))
				v_premi_operator = v_premi_operator + cdbl(to_nol(rsGrnInv("premi_operator")))

				rsGrnInv.MoveNext
			Wend
		rsGrnInv.Close
		set rsGrnInv = Nothing
	%>
			<tr bgcolor="#F4F4A0"> 
				<td align="center" colspan="5"><font style="font-weight: bold;">TOTAL</font></td>				
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_hasilkerja)),2)%></font></td> 
				<td align="right"></td>
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_lebih_basis)),2)%></font></td>  
				<td align="right"></td>
				<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_premi_operator)),2)%></font></td>  
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
	
</script>