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

	v_txtFcba = Request("txtFcba")
	v_txtMonth = Request("txtMonth")
	v_txtYear = Request("txtYear")
	v_txtsectionname = Request("txtsectionname")
	v_txtempcode = Request("txtempcode")

	txttdate = txtMonth & "" & txtYear
	If "" & v_txtYear = "" Then v_txtYear = Year(Session("v_Costbook"))
	If "" & v_txtMonth = "" Then v_txtMonth = Month(Session("v_Costbook"))
	If "" & txttdate = "" Then txttdate = textDate(date)
	varDefaultBa = Session("fcBA")

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
	 
		strGrnInv	=	"   SELECT fcba," & _
						"		   fccode, " & _
         				"		   employee_api.get_employeename (fcba, fccode)employeename, "  & _
        				"		   employee_api.get_sectionname(fcba, fccode) sectionname, " & _
						"          hdb_month," & _
						"          hdb_year," & _
						"          FDDATE," & _ 
						"          MAX (lebih_basis_1) lebih_basis_1," & _
						"          MAX (lebih_basis_2) lebih_basis_2," & _
						"          MAX (presentrate) presentrate," & _
						"          MAX (insentifrate) insentifrate," & _
						"          SUM (extrainsentifrate01_basis_1) extrainsentifrate01_basis_1," & _
						"          SUM (insentif_1_1) insentif_1_1," & _
						"          SUM (insentif_1_2) insentif_1_2," & _
						"          SUM (insentif_2_1) insentif_2_1" & _
						"     FROM (SELECT fcba," & _
						"                  hdb_month," & _
						"                  hdb_year," & _
						"                  LAST_DAY (TO_DATE (hdb_month || hdb_year, 'mmyyyy')) FDDATE," & _
						"                  employeecode FCCODE," & _
						"                  lebih_basis_1," & _
						"                  lebih_basis_2," & _
						"                  presentrate," & _
						"                  insentifrate," & _
						"                  extrainsentifrate01_basis_1," & _
						"                  lebih_basis_1 * insentifrate AS insentif_1_1," & _
						"                  extrainsentifrate01_basis_1 * presentrate AS insentif_1_2," & _
						"                  insentif_2_1" & _
						"             FROM (SELECT hdbb.fcba," & _
						"                          hdbb.hdb_month AS hdb_month," & _
						"                          hdbb.hdb_year AS hdb_year," & _
						"                          hdbb.employeecode AS employeecode," & _
						"                          hdbb.lebih_basis_1 AS lebih_basis_1," & _
						"                          hdbb.lebih_basis_2 AS lebih_basis_2," & _
						"                          emp.presentrate AS presentrate," & _
						"                          CASE" & _
						"                             WHEN hdbb.lebih_basis_1 BETWEEN tmpll.MINHK" & _
						"                                                         AND tmpll.MAXHK" & _
						"                             THEN" & _
						"                                tmpll.insentifrate" & _
						"                             ELSE 0" & _
						"                          END" & _
						"                             AS insentifrate," & _
						"                          CASE" & _
						"                             WHEN hdbb.lebih_basis_1 BETWEEN tmpll.MINHK" & _
						"                                                         AND tmpll.MAXHK" & _
						"                             THEN" & _
						"                                tmpll.extrainsentifrate01_basis_1" & _
						"                             ELSE 0" & _
						"                          END" & _
						"                             AS extrainsentifrate01_basis_1," & _
						"                          CASE" & _
						"                             WHEN hdbb.lebih_basis_2 BETWEEN tmpll.MINHK" & _
						"                                                         AND tmpll.MAXHK" & _
						"                             THEN" & _
						"                                tmpll.extrainsentifrate02_basis_2" & _
						"                             ELSE 0" & _
						"                          END" & _
						"                             AS insentif_2_1" & _
						"                     FROM (  SELECT fcba," & _
						"                                    hdb_month," & _
						"                                    hdb_year," & _
						"                                    employeecode," & _
						"                                    SUM (lebih_basis_1) AS lebih_basis_1," & _
						"                                    SUM (lebih_basis_2) AS lebih_basis_2" & _
						"                               FROM (  SELECT fcba," & _
						"                                              fddate," & _
						"                                              TO_CHAR (" & _
						"                                                 TO_DATE (FDDATE, 'MM/DD/YYYY')," & _
						"                                                 'mm')" & _
						"                                                 AS hdb_month," & _
						"                                              TO_CHAR (" & _
						"                                                 TO_DATE (FDDATE, 'MM/DD/YYYY')," & _
						"                                                 'yyyy')" & _
						"                                                 AS hdb_year," & _
						"                                              employeecode," & _
						"                                              SUM (quantity) AS quantity," & _
						"                                              AVG (basis_1) AS basis_1," & _
						"                                              AVG (basis_2) AS basis_2," & _
						"                                              CASE" & _
						"                                                 WHEN SUM (quantity) > AVG (basis_1)" & _
						"                                                 THEN 1" & _
						"                                                 ELSE 0" & _
						"                                              END" & _
						"                                                 AS lebih_basis_1," & _
						"                                              CASE" & _
						"                                                 WHEN SUM (quantity) > AVG (basis_2)" & _
						"                                                 THEN 1" & _
						"                                                 ELSE 0" & _
						"                                              END" & _
						"                                                 AS lebih_basis_2" & _
						"                                         FROM (SELECT hdb.fcba," & _
						"                                                      hdb.fddate," & _
						"                                                      hdb.employeecode," & _
						"                                                      hdb.division," & _
						"                                                      hdb.plantingyear," & _
						"                                                      hdb.presentgrade," & _
						"                                                      hdb.quantity," & _
						"                                                      tmpl.basis_1," & _
						"                                                      tmpl.basis_2" & _
						"                                                 FROM (  SELECT hdbj.fcba," & _
						"                                                                hdbj.fddate," & _
						"                                                                hdbj.employeecode," & _
						"                                                                hdbj.division," & _
						"                                                                hdbj.plantingyear," & _
						"                                                                hdbj.presentgrade," & _
						"                                                                CASE" & _
						"                                                                   WHEN MAX (" & _
						"                                                                           ald.type_1) =" & _
						"                                                                           'D'" & _
						"                                                                   THEN" & _
						"                                                                      SUM (" & _
						"                                                                         hdbj.quantity_bydaily)" & _
						"                                                                   ELSE" & _
						"                                                                      SUM (" & _
						"                                                                         hdbj.quantity_bymonthly)" & _
						"                                                                END" & _
						"                                                                   AS quantity" & _
						"                                                           FROM harvesting_daily_bjr" & _
						"                                                                hdbj" & _
						"                                                                INNER JOIN" & _
						"                                                                (SELECT allowded.fcba," & _
						"                                                                        allowded.fccode," & _
						"                                                                        allowded.type_1" & _
						"                                                                   FROM allowded" & _
						"                                                                  WHERE     allowded.fccode =" & _
						"                                                                               'A27'" & _
						"                                                                        AND allowded.fcba =" & _
						"                                                                               '"& v_txtfcba &"')" & _
						"                                                                ald" & _
						"                                                                   ON (hdbj.fcba =" & _
						"                                                                          ald.fcba)" & _
						"                                                          WHERE     hdbj.fcba = '"& v_txtfcba &"'" & _
						"                                                                AND TO_NUMBER (" & _
						"                                                                       TO_CHAR (" & _
						"                                                                          TO_DATE (" & _
						"                                                                             hdbj.fddate," & _
						"                                                                             'MM/DD/YYYY')," & _
						"                                                                          'mm')) = '"& v_txtmonth &"'" & _
						"                                                                AND TO_NUMBER (" & _
						"                                                                       TO_CHAR (" & _
						"                                                                          TO_DATE (" & _
						"                                                                             hdbj.fddate," & _
						"                                                                             'MM/DD/YYYY')," & _
						"                                                                          'yyyy')) = '"& v_txtyear &"'" & _
						"                                                       GROUP BY hdbj.fcba," & _
						"                                                                hdbj.fddate," & _
						"                                                                hdbj.employeecode," & _
						"                                                                hdbj.division," & _
						"                                                                hdbj.plantingyear," & _
						"                                                                hdbj.presentgrade)" & _
						"                                                      hdb" & _
						"                                                      INNER JOIN" & _
						"                                                      (  SELECT harvesting_premiinsentiftmpl.fcba," & _
						"                                                                harvesting_premiinsentiftmpl.division," & _
						"                                                                harvesting_premiinsentiftmpl.plantingyear," & _
						"                                                                harvesting_premiinsentiftmpl.employeetype," & _
						"                                                                MAX (" & _
						"                                                                   harvesting_premiinsentiftmpl.basis_1)" & _
						"                                                                   AS basis_1," & _
						"                                                                MAX (" & _
						"                                                                   harvesting_premiinsentiftmpl.basis_2)" & _
						"                                                                   AS basis_2" & _
						"                                                           FROM HARVESTING_PREMIINSENTIFTMPL" & _
						"                                                          WHERE HARVESTING_PREMIINSENTIFTMPL.FCBA =" & _
						"                                                                   '"& v_txtfcba &"'" & _
						"                                                       GROUP BY harvesting_premiinsentiftmpl.fcba," & _
						"                                                                harvesting_premiinsentiftmpl.division," & _
						"                                                                harvesting_premiinsentiftmpl.plantingyear," & _
						"                                                                harvesting_premiinsentiftmpl.employeetype)" & _
						"                                                      tmpl" & _
						"                                                         ON (    hdb.fcba =" & _
						"                                                                    tmpl.fcba" & _
						"                                                             AND hdb.division =" & _
						"                                                                    tmpl.division" & _
						"                                                             AND hdb.plantingyear =" & _
						"                                                                    tmpl.plantingyear" & _
						"                                                             AND hdb.presentgrade =" & _
						"                                                                    tmpl.employeetype))" & _
						"                                        WHERE     fcba = '"& v_txtfcba &"'" & _
						"                                              AND TO_NUMBER(TO_CHAR (" & _
						"                                                     TO_DATE (FDDATE, 'MM/DD/YYYY')," & _
						"                                                     'mm')) = '"& v_txtmonth &"'" & _
						"                                              AND TO_NUMBER(TO_CHAR (" & _
						"                                                     TO_DATE (FDDATE, 'MM/DD/YYYY')," & _
						"                                                     'yyyy')) = '"& v_txtYear &"'" & _ 
						"                                     GROUP BY fcba, fddate, employeecode)" & _
						"                              WHERE     fcba = '"& v_txtfcba &"'" & _
						"                                    AND TO_NUMBER(TO_CHAR (TO_DATE (FDDATE, 'MM/DD/YYYY')," & _
						"                                                 'mm')) = '"& v_txtMonth &"'" & _
						"                                    AND TO_NUMBER(TO_CHAR (TO_DATE (FDDATE, 'MM/DD/YYYY')," & _
						"                                                 'yyyy')) = '"& v_txtYear &"'" & _ 
						"                           GROUP BY fcba," & _
						"                                    hdb_month," & _
						"                                    hdb_year," & _
						"                                    employeecode) hdbb" & _
						"                          INNER JOIN" & _
						"                          (  SELECT harvesting_premiinsentiftmpl.fcba," & _
						"                                    harvesting_premiinsentiftmpl.minhk," & _
						"                                    harvesting_premiinsentiftmpl.maxhk," & _
						"                                    MAX (" & _
						"                                       harvesting_premiinsentiftmpl.insentifrate)" & _
						"                                       AS insentifrate," & _
						"                                    MAX (" & _
						"                                       harvesting_premiinsentiftmpl.extrainsentifrate01_basis_1)" & _
						"                                       AS extrainsentifrate01_basis_1," & _
						"                                    MAX (" & _
						"                                       harvesting_premiinsentiftmpl.extrainsentifrate02_basis_2)" & _
						"                                       extrainsentifrate02_basis_2" & _
						"                               FROM HARVESTING_PREMIINSENTIFTMPL" & _
						"                              WHERE HARVESTING_PREMIINSENTIFTMPL.FCBA = '"& v_txtfcba &"'" & _
						"                           GROUP BY harvesting_premiinsentiftmpl.fcba," & _
						"                                    harvesting_premiinsentiftmpl.minhk," & _
						"                                    harvesting_premiinsentiftmpl.maxhk) tmpll" & _
						"                             ON hdbb.fcba = tmpll.fcba" & _
						"                          INNER JOIN" & _
						"                          (SELECT employee.fcba," & _
						"                                  employee.fccode," & _
						"                                  employee.presentgrade," & _
						"                                  CASE" & _
						"                                     WHEN employee.presentgrade = 'KHT'" & _
						"                                     THEN" & _
						"                                        employee.presentrate / 30" & _
						"                                     ELSE" & _
						"                                        employee.presentrate" & _
						"                                  END" & _
						"                                     AS presentrate" & _
						"                             FROM employee" & _
						"                            WHERE employee.emp_type = 'PEMANEN') emp" & _
						"                             ON (    hdbb.fcba = emp.fcba" & _
						"                                 AND hdbb.employeecode = emp.fccode)" & _
						"                    WHERE     hdbb.fcba = '"& v_txtfcba &"'" & _
						"                          AND TO_NUMBER (hdbb.hdb_year) = '"& v_txtYear &"'" & _
						"                          AND TO_NUMBER (hdbb.hdb_month) = '"& v_txtMonth &"')" & _
						"            WHERE fcba = '"& v_txtfcba &"'" 
   
	if  len(trim(v_txtempcode)) > 0 then
		strGrnInv = strGrnInv & " AND employeecode ='"& v_txtempcode &"'"
	end if 
	if  len(trim(v_txtsectionname)) > 0 then
		strGrnInv = strGrnInv & " AND employee_api.get_sectionname(fcba, employeecode) ='"& v_txtsectionname &"'"
	end if 
	if  len(trim(v_txtYear)) > 0 then
		strGrnInv = strGrnInv & " AND hdb_year ="& v_txtYear &""
	end if 
	if  len(trim(v_txtmonth)) > 0 then
		strGrnInv = strGrnInv & " AND hdb_month ="& v_txtmonth &")  WHERE (insentif_1_1 + insentif_1_2 + insentif_2_1) > 0"
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

	strGrnInv	=	strGrnInv & " GROUP BY fcba," & _
								"          hdb_month," & _
								"          hdb_year," & _
								"          FDDATE," & _
								"          FCCODE"  
								
	strGrnInv	=	strGrnInv & " ORDER BY  fcba," & _
								"        hdb_month," & _
								"        hdb_year," & _
								"        fccode"  

	' response.write strGrnInv	
	Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
	rsGrnInv.Open strGrnInv, connstr, adOpenStatic 
	
%>


<html>

<head>
<title>ePlaS Reporting - Premi Intensif</title>
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
							<font face="verdana" size="1" style="font-weight: bold;">TAHUN</font>
						</td>
						<td align="center">
							<font face="verdana" size="1" style="font-weight: bold;">BULAN</font>
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
									if v_txtFcba  = rsFcba("fcba") then
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
							<select name="txtempcode" class="chosen-select" >
								<option value="">--ALL--</option>
									<% 
									if len(trim(v_txtempcode)) > 0 Then
										strsectionname	=	" select fccode, fcname from employee where fcba = '"& v_txtfcba &"' order by fcCode"
									Else
										strsectionname	=	" select fccode, fcname from employee where fcba = '"& Session("fcba") &"' order by fcCode"
									End If
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
    <th colspan="8" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
				<h4>PT. <%=CompanyName%></h4>
				<h5><%=fcbaname%></h5>
			</div>
			
			<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
				<h2>LAPORAN PREMI INSENTIF</h2>
				<h4>PERIODE : Year <%=v_txtyear%>, Month <%=v_txtmonth%></h4>
                <br/>
			</div>
		</div>
</th>
		</tr>
		<tr>
			<th style="background-color:#00A300;color:#fff;">Employee Code</th>
			<th style="background-color:#00A300;color:#fff;">Employee Name</th>
			<th style="background-color:#00A300;color:#fff;">Lebih Basis 1</th>
			<th style="background-color:#00A300;color:#fff;">Lebih Basis 2</th> 
			<th style="background-color:#00A300;color:#fff;">Insentif 1</th> 
			<th style="background-color:#00A300;color:#fff;">Insentif 2</th> 
			<th style="background-color:#00A300;color:#fff;">Insentif 3</th> 
			<th style="background-color:#00A300;color:#fff;">Total Insentif</th>  
		</tr>
    </thead>
	<tbody>
<%if rsGrnInv.BOF and rsGrnInv.EOF then%>
		<tr>
			<td colspan="8" align="left">NO DATA AVAILABLE</td>
		</tr>
<%end if
		NO = 1
		v_lebih_basis_1 = 0
		v_lebih_basis_2 = 0 
		v_insentif_1_1 = 0
		v_insentif_1_2 = 0
		v_insentif_2_1 = 0
		v_totalinsentif = 0 
		while not rsGrnInv.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
	    totalinsentif = cdbl(to_nol(rsGrnInv("insentif_1_1"))) + cdbl(to_nol(rsGrnInv("insentif_1_2"))) + cdbl(to_nol(rsGrnInv("insentif_2_1")))
%>
		<tr class="<%=cClass%>">
			<td align="center"><%=rsGrnInv("fccode")%></td>
			<td align="Left"><%=rsGrnInv("employeename")%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("lebih_basis_1"))),2)%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("lebih_basis_2"))),2)%></td> 
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("insentif_1_1"))),2)%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("insentif_1_2"))),2)%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("insentif_2_1"))),2)%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(totalinsentif)),2)%></td>
		</tr>
<%			
			v_lebih_basis_1 = v_lebih_basis_1 + cdbl(to_nol(rsGrnInv("lebih_basis_1"))) 
			v_lebih_basis_2 = v_lebih_basis_2 + cdbl(to_nol(rsGrnInv("lebih_basis_2")))  
			v_insentif_1_1 = v_insentif_1_1 + cdbl(to_nol(rsGrnInv("insentif_1_1"))) 
			v_insentif_1_2 = v_insentif_1_2 + cdbl(to_nol(rsGrnInv("insentif_1_2"))) 
			v_insentif_2_1 = v_insentif_2_1 + cdbl(to_nol(rsGrnInv("insentif_2_1"))) 
			v_totalinsentif = v_totalinsentif + cdbl(to_nol(totalinsentif)) 
			rsGrnInv.MoveNext
		Wend
	rsGrnInv.Close
	set rsGrnInv = Nothing
%>
		<tr bgcolor="#F4F4A0"> 
			<td align="center" colspan="2"><font style="font-weight: bold;">TOTAL</font></td>				
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_lebih_basis_1)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_lebih_basis_2)),2)%></font></td> 
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_insentif_1_1)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_insentif_1_2)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_insentif_2_1)),2)%></font></td>
			<td align="right"><font style="font-weight: bold;"><%=FormatNumber(cdbl(to_nol(v_totalinsentif)),2)%></font></td>
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