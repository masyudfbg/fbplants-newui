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
	If "" & txtYear = "" Then txtYear = Year(Session("v_Costbook"))
	If "" & txtMonth = "" Then txtMonth = Month(Session("v_Costbook"))
	If "" & txttdate = "" Then txttdate = textDate(date)
	varDefaultBa = Session("fcBA")

	v_txtdatefrom = Request("txtdatefrom")
	v_txtdateuntil = Request("txtdateuntil")
	If (v_txtdatefrom = "" Or Len(Trim(v_txtdatefrom)) = 0) Then v_txtdatefrom = textdate(date)
	If (v_txtdateuntil = "" Or Len(Trim(v_txtdateuntil)) = 0) Then v_txtdateuntil = textdate(date)

	If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
		v_txtFcba = varDefaultBa
	End If	
	nSpanTotal = 0
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
	
	q = "select type_1 from allowded where fccode = 'A26' and fcba ='" & v_txtfcba & "' "
	set rsCat = connStr.Execute(q)
	if not rsCat.eof then
	   type_1 = rsCat("type_1")
	end if
	
	If type_1 = "M" Then 
		strGrnInv 	= 	"   SELECT ag.FCBA," &  _ 
						"          ag.FDDATE," &  _ 
						"          ag.EMPLOYEECODE FCCODE," &  _ 
						"          employee_api.get_employeename (ag.FCBA,  ag.EMPLOYEECODE) employeename," &  _ 
						"          ff.DIVISION," &  _ 
						"          SUM (NVL (ag.BUCKET, 0)) AS bucket," &  _ 
						"          MAX (BBR.BBR) AS bbr," &  _ 
						"          (SUM (NVL (ag.BUCKET, 0)) * MAX (BBR.BBR)) AS totalkg," &  _ 
						"          SUM (NVL (ag.BUCKET, 0)) * MAX (RATE.RATE_BRONDOLAN) * MAX (BBR.BBR)" &  _ 
						"             AMOUNT" &  _ 
						"     FROM attendance_gad ag" &  _ 
						"          INNER JOIN field ff" &  _ 
						"             ON     ff.fccode = ag.LOCATIONCODE" &  _ 
						"                AND ff.FCBA = ag.FCBA" &  _ 
						"                AND ff.ACTIVATION = 'Y'" &  _ 
						"          INNER JOIN (  SELECT FCBA, DIVISION, MAX (RATE_BRONDOLAN) RATE_BRONDOLAN" &  _ 
						"                          FROM MULTILEVEL_HARVESTING_RATE" &  _ 
						"                         WHERE fcba = '"& v_txtFcba &"'" &  _ 
						"                      GROUP BY FCBA, DIVISION) RATE" &  _ 
						"             ON RATE.fcba = ag.fcba AND ff.DIVISION = RATE.DIVISION" &  _ 
						"          INNER JOIN" &  _ 
						"          (  SELECT FCBA," &  _ 
						"                    TO_CHAR (HARVESTDATE, 'YYYY') YEARSX," &  _ 
						"                    TO_CHAR (HARVESTDATE, 'mm') MONTHX," &  _ 
						"                    DIVISION," &  _ 
						"                    SUM (KGPERBLOK) AS KGPERBLOK," &  _ 
						"                    SUM (bucket) AS bucket," &  _ 
						"                    SUM (KGPERBLOK) / SUM (BUCKET) BBR" &  _ 
						"               FROM (SELECT spb.FCBA," &  _ 
						"                            spb.harvestdate AS harvestdate," &  _ 
						"                            spb.SPBNO," &  _ 
						"                            spb.FIELDCODE," &  _ 
						"                            ff.DIVISION," &  _ 
						"                            spb.BUCKET," &  _ 
						"                            wh.total_bucket," &  _ 
						"                            wh.Kg_Brondolan," &  _ 
						"                            spb.BUCKET / wh.total_bucket * wh.Kg_Brondolan" &  _ 
						"                               KGPERBLOK" &  _ 
						"                       FROM wb_harvesting wh" &  _ 
						"                            INNER JOIN harvestingspb spb" &  _ 
						"                               ON     wh.SPBNO = spb.SPBNO" &  _ 
						"                                  AND wh.FCBA_ESTATE = spb.FCBA" &  _ 
						"                                  AND spb.FCBA = '"& v_txtFcba &"'" &  _ 
						"                            INNER JOIN field ff" &  _ 
						"                               ON     ff.fccode = spb.FIELDCODE" &  _ 
						"                                  AND ff.FCBA = spb.FCBA" &  _ 
						"                                  AND ff.ACTIVATION = 'Y'" &  _ 
						"                      WHERE     fcba_estate = '"& v_txtFcba &"'" &  _ 
						"                            AND spb.HARVESTDATE >= TO_DATE('" & v_txtdatefrom & "', 'mm/dd/yyyy')" &  _ 
						"                            AND spb.HARVESTDATE <= TO_DATE('" & v_txtdateuntil & "', 'mm/dd/yyyy')" &  _  
						"                            AND wh.TBSTYPE = 'TBSB')" &  _ 
						"              WHERE     fcba = '"& v_txtFcba &"'" &  _ 
						"                    AND HARVESTDATE >= TO_DATE('" & v_txtdatefrom & "', 'mm/dd/yyyy')" &  _ 
						"                    AND HARVESTDATE <= TO_DATE('" & v_txtdateuntil & "', 'mm/dd/yyyy')" &  _  
						"           GROUP BY FCBA," &  _ 
						"                    TO_CHAR (HARVESTDATE, 'YYYY')," &  _ 
						"                    TO_CHAR (HARVESTDATE, 'mm')," &  _ 
						"                    DIVISION) BBR" &  _ 
						"             ON (    BBR.fcba = ag.fcba" &  _ 
						"                 AND ff.DIVISION = BBR.DIVISION" &  _ 
						"                 AND TO_CHAR (ag.FDDATE, 'yyyy') = BBR.YEARSX" &  _ 
						"                 AND TO_CHAR (ag.FDDATE, 'mm') = BBR.MONTHX)" &  _ 
						"    WHERE     ag.fcba = '"& v_txtFcba &"'" &  _  
						"          AND NVL (ag.BUCKET, 0) > 0" &  _ 
						"          AND ag.ATTENDANCE IN ('KJ'," &  _ 
						"                                'WH'," &  _ 
						"                                'WS'," &  _ 
						"                                'KK')"  
											
						
			if  len(trim(v_txtdatefrom)) > 0 then
				strGrnInv = strGrnInv & " AND ag.FDDATE >= TO_DATE('" & v_txtdatefrom & "', 'mm/dd/yyyy')"
				txtSelect = " From " & FormatDateTime(v_txtdatefrom,1)
			end if
			if  len(trim(v_txtdateuntil)) > 0 then
				strGrnInv = strGrnInv & " AND ag.FDDATE <= TO_DATE('" & v_txtdateuntil & "', 'mm/dd/yyyy')"
				txtSelect = txtSelect & " To " & FormatDateTime(v_txtdateuntil,1)
			end if 
			if  len(trim(v_txtsectionname)) > 0 then
				strGrnInv	=	strGrnInv & " AND ff.division = '"& v_txtsectionname &"'" 
				txtSelect = txtSelect & " Division Code " & v_txtsectionname
			end if
			if  len(trim(v_txtempcode)) > 0 then
				strGrnInv	=	strGrnInv & " AND FCCODE = '"& v_txtempcode &"'" 
				txtSelect = txtSelect & " Employee Code " & v_txtempcode
			end if 
			
			strGrnInv   =   strGrnInv & " GROUP BY ag.FCBA," &  _ 
										"          ag.FDDATE," &  _ 
										"          ag.EMPLOYEECODE," &  _ 
										"          ff.DIVISION" 	
										
			strGrnInv	=	strGrnInv & " ORDER BY  ag.FCBA," &  _ 
										"          ag.FDDATE," &  _ 
										"          ag.EMPLOYEECODE," &  _ 
										"          ff.DIVISION" 	

	Else
		strGrnInv	=	" SELECT emp.fcba," & _
						"        emp.FCCODE," & _
						"        EMPLOYEE_API.GET_EMPLOYEENAME (emp.fcba, emp.FCCODE) employeename," & _
						"        mwn.harvestdate," & _
						"        att_gad.LOCATIONCODE," & _
						"        ff.division," & _
						"        att_gad.BUCKET," & _
						"        ROUND (MWN.BBR, 2) BBR," & _
						"        (BBR * att_gad.BUCKET) totalkg," & _
						"        ROUND ( (att_gad.bucket * MWN.BBR * mhr.rate_brondolan), 2) AS amount" & _
						"   FROM attendance_gad att_gad" & _
						"        INNER JOIN Employee Emp" & _
						"           ON     att_gad.EMPLOYEECODE = emp.FCCODE" & _
						"              AND att_gad.FCBA = emp.FCBA" & _
						"              AND emp.DATETERMINATE IS NULL" & _
						"        INNER JOIN" & _
						"        (SELECT FCBA_ESTATE," & _
						"                harvestdate," & _
						"                DIVISION," & _
						"                fieldcode," & _
						"                total_bunch AS bunch," & _
						"                kg_per_block," & _
						"                Kg_Brondolan," & _
						"                bucket_per_block," & _
						"                total_bucket," & _
						"                BBR" & _
						"           FROM (  SELECT FCBA_ESTATE," & _
						"                          harvestdate," & _
						"                          DIVISION," & _
						"                          fieldcode," & _
						"                          SUM (bunch) AS bunch," & _
						"                          SUM (total_bunch) AS total_bunch," & _
						"                          SUM (Bunch_Mill_Weight_Netto)" & _
						"                             AS Bunch_Mill_Weight_Netto," & _
						"                          SUM (kg_per_block) AS kg_per_block," & _
						"                          SUM (Kg_Brondolan) AS Kg_Brondolan," & _
						"                          SUM (bucket_per_block) AS bucket_per_block," & _
						"                          SUM (total_bucket) AS total_bucket," & _
						"                          NVL (" & _
						"                             DECODE (SUM (total_bucket)," & _
						"                                     0, 0," & _
						"                                     SUM (Kg_Brondolan) / SUM (total_bucket))," & _
						"                             0)" & _
						"                             AS BBR" & _
						"                     FROM (SELECT iplasprod.harvestingspb.fcba AS FCBA_ESTATE," & _
						"                                  iplasprod.harvestingspb.spbno," & _
						"                                  iplasprod.harvestingspb.harvestdate," & _
						"                                  iplasprod.harvestingspb.fieldcode," & _
						"                                  FIELD.DIVISION," & _
						"                                  NVL (iplasprod.harvestingspb.bunch, 0)" & _
						"                                     AS bunch," & _
						"                                  NVL (total_bunch.total_bunch, 0)" & _
						"                                     AS total_bunch," & _
						"                                  NVL (wb_harv.Bunch_Mill_Weight_Netto, 0)" & _
						"                                     AS Bunch_Mill_Weight_Netto," & _
						"                                  DECODE (" & _
						"                                     total_bunch.total_bunch," & _
						"                                     0, 0," & _
						"                                     (  (  iplasprod.harvestingspb.bunch" & _
						"                                         / total_bunch.total_bunch)" & _
						"                                      * wb_harv.Bunch_Mill_Weight_Netto))" & _
						"                                     AS kg_per_block," & _
						"                                  Kg_Brondolan," & _
						"                                  NVL (iplasprod.harvestingspb.bucket, 0)" & _
						"                                     AS bucket_per_block," & _
						"                                  total_bucket" & _
						"                             FROM iplasprod.harvestingspb" & _
						"                                  INNER JOIN FIELD" & _
						"                                     ON     iplasprod.harvestingspb.fieldcode =" & _
						"                                               field.fccode" & _
						"                                        AND iplasprod.harvestingspb.fcba =" & _
						"                                               field.fcba" & _
						"                                        AND field.ACTIVATION = 'Y'" & _
						"                                  INNER JOIN" & _
						"                                  (  SELECT fcba," & _
						"                                            spbno," & _
						"                                            SUM (bunch) AS total_bunch" & _
						"                                       FROM iplasprod.harvestingspb" & _
						"                                   GROUP BY fcba, spbno) total_bunch" & _
						"                                     ON (    iplasprod.harvestingspb.fcba =" & _
						"                                                total_bunch.fcba" & _
						"                                         AND iplasprod.harvestingspb.spbno =" & _
						"                                                total_bunch.spbno)" & _
						"                                  INNER JOIN" & _
						"                                  (  SELECT iplasprod.wb_harvesting.fcba_estate," & _
						"                                            iplasprod.wb_harvesting.spbno," & _
						"                                            SUM (" & _
						"                                               iplasprod.wb_harvesting.bunch_mill_weight_netto)" & _
						"                                               AS Bunch_Mill_Weight_Netto," & _
						"                                            SUM (" & _
						"                                               iplasprod.wb_harvesting.Kg_Brondolan)" & _
						"                                               Kg_Brondolan," & _
						"                                            SUM (" & _
						"                                               iplasprod.wb_harvesting.total_bucket)" & _
						"                                               total_bucket" & _
						"                                       FROM iplasprod.wb_harvesting" & _
						"                                   GROUP BY iplasprod.wb_harvesting.fcba_estate," & _
						"                                            iplasprod.wb_harvesting.spbno)" & _
						"                                  wb_harv" & _
						"                                     ON (    iplasprod.harvestingspb.fcba =" & _
						"                                                wb_harv.fcba_estate" & _
						"                                         AND iplasprod.harvestingspb.spbno =" & _
						"                                                wb_harv.spbno))" & _
						"                 GROUP BY FCBA_ESTATE," & _
						"                          harvestdate," & _
						"                          DIVISION," & _
						"                          fieldcode)) MWN" & _
						"           ON     (    att_gad.fcba = MWN.FCBA_ESTATE" & _
						"                   AND att_gad.LOCATIONCODE = MWN.FIELDCODE)" & _
						"              AND TO_CHAR (att_gad.FDDATE, 'MM/DD/YYYY') =" & _
						"                     TO_CHAR (MWN.harvestdate, 'MM/DD/YYYY')" & _
						"        INNER JOIN FIELD ff" & _
						"           ON     att_gad.LOCATIONCODE = ff.FCCODE" & _
						"              AND att_gad.fcba = ff.fcba" & _
						"              AND ff.ACTIVATION = 'Y'" & _
						"        INNER JOIN (SELECT DISTINCT fcba," & _
						"                                    division," & _
						"                                    planting_year," & _
						"                                    rate_brondolan" & _
						"                      FROM MULTILEVEL_HARVESTING_RATE" & _
						"                     WHERE fcba = '"& v_txtfcba &"') mhr" & _
						"           ON (    att_gad.fcba = mhr.fcba" & _
						"               AND ff.DIVISION = mhr.DIVISION" & _
						"               AND TO_CHAR (ff.PLANTINGDATE, 'yyyy') = mhr.PLANTING_YEAR)" & _
						"  WHERE     att_gad.Attendance IN ('KJ'," & _
						"                                   'WH'," & _
						"                                   'WS'," & _
						"                                   'KK')" & _
						"        AND att_gad.Jobcode IN ('112030101', '505030101')" & _
						"        AND att_gad.LOCATIONTYPE = 'FF'" & _
						"        AND att_gad.MANDAYS > 0" & _
						"        AND att_gad.BUCKET > 0" & _
						"        AND att_gad.rate = 0" & _
						"        AND att_gad.fcBA = '"& v_txtfcba &"'"  
		
			if  len(trim(v_txtdatefrom)) > 0 then
				strGrnInv = strGrnInv & " AND MWN.harvestdate >= TO_DATE('" & v_txtdatefrom & "', 'mm/dd/yyyy')"
				txtSelect = " From " & FormatDateTime(v_txtdatefrom,1)
			end if
			if  len(trim(v_txtdateuntil)) > 0 then
				strGrnInv = strGrnInv & " AND MWN.harvestdate <= TO_DATE('" & v_txtdateuntil & "', 'mm/dd/yyyy')"
				txtSelect = txtSelect & " To " & FormatDateTime(v_txtdateuntil,1)
			end if 
			if  len(trim(v_txtsectionname)) > 0 then
				strGrnInv	=	strGrnInv & " AND ff.division = '"& v_txtsectionname &"'" 
				txtSelect = txtSelect & " Division Code " & v_txtsectionname
			end if
			if  len(trim(v_txtempcode)) > 0 then
				strGrnInv	=	strGrnInv & " AND emp.FCCODE = '"& v_txtempcode &"'" 
				txtSelect = txtSelect & " Employee Code " & v_txtempcode
			end if 
		
		
			
			strGrnInv	=	strGrnInv & " ORDER BY fcba," & _
										"          FCCODE," & _ 
										"          harvestdate," & _
										" 		   LOCATIONCODE" 
		End If 
		' response.write strGrnInv	
		Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
		rsGrnInv.Open strGrnInv, connstr, adOpenStatic 
 
%>


<html>

<head>
<title>ePlaS Reporting - Premi Brondolan </title>
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
						<td>
							<font face="verdana" size="1">KEBUN</font>
						</td>
						<td>
							<select name="txtFcba" id="vFcba" onchange="javascript:filtering()">
							<%
							Dim rsFcba
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
						<td>
							<font face="verdana" size="1">FROM</font>
						</td>
						<td width="4">
							<input class="datepicker" type="text" maxlength="10" name="txtdatefrom" size="10" value="<%=v_txtdatefrom%>">
						</td>
						<td>
							<font face="verdana" size="1">TO</font>
						</td>
						<td width="2">
							<input class="datepicker" type="text" maxlength="10" name="txtdateuntil" size="10" value="<%=v_txtdateuntil%>">
						</td>
						<td>
							<font face="verdana" size="1">DIVISION</font>
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
							<font face="verdana" size="1">EMPLOYEE</font>
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
    		<th colspan="9" style="background-color:#FFFFFF; border:none;">  
				<div id="header-wrapper">
					<div id="head_judul" style="float:left; text-align:left;">
						<h4>PT. <%=CompanyName%></h4>
						<h5><%=fcbaname%></h5>
					</div>
					<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
						<h2>LAPORAN PREMI BRONDOLAN</h2>
		                <br/>
						<h4>Selected By :<%=txtSelect%></h4>
					</div> 
				</div>
			</th>
		</tr>
		<tr>
			<th style="background-color:#00A300;color:#fff;">EMPLOYEE CODE</th> 
			<th style="background-color:#00A300;color:#fff;">EMPLOYEE NAME</th>
			<th style="background-color:#00A300;color:#fff;">DIVISION</th>
			<th style="background-color:#00A300;color:#fff;">DATE</th>
			<% if type_1 = "D" Then%>
			<th style="background-color:#00A300;color:#fff;">BLOCK</th>
			<%end if%> 
			<th style="background-color:#00A300;color:#fff;">BUCKET</th>
			<th style="background-color:#00A300;color:#fff;">BBR</th>
			<th style="background-color:#00A300;color:#fff;">KG BRONDOLAN</th>
			<th style="background-color:#00A300;color:#fff;">AMOUNT</th>
		</tr>   
    </thead>
<tbody>
<%if rsGrnInv.BOF and rsGrnInv.EOF then%>
		<tr>
			<td colspan="9" align="left">NO DATA AVAILABLE</td>
		</tr>
<%end if
		NO = 1
		rows = 1
		nbucket = 0
		nbbr = 0
		namount = 0
		ntotalkg = 0
		while not rsGrnInv.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
		<tr class="<%=cClass%>"> 
			<td align="center"><%=rsGrnInv("fccode")%></td>
			<td align="left"><%=rsGrnInv("employeename")%></td> 
			<td align="center"><%=rsGrnInv("division")%></td> 
			<% if type_1 = "M" Then%>
				<td align="center"><%=rsGrnInv("fddate")%></td>
			<%else%>
				<td align="center"><%=rsGrnInv("harvestdate")%></td>
				<td align="center"><%=rsGrnInv("LOCATIONCODE")%></td>
			<%end if%> 
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("bucket"))),0)%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("bbr"))),2)%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("totalkg"))),2)%></td>
			<td align="right"><%=FormatNumber(cdbl(to_nol(rsGrnInv("amount"))),2)%></td>  
		</tr> 
	 
		<%  
			nbucket = nbucket + cdbl(to_nol(rsGrnInv("bucket")))
			nbbr = nbbr + cdbl(to_nol(rsGrnInv("bbr")))	
			namount = namount + cdbl(to_nol(rsGrnInv("amount")))  
			ntotalkg = ntotalkg + cdbl(to_nol(rsGrnInv("totalkg")))  
			rsGrnInv.MoveNext			
			Wend
		rsGrnInv.Close
		set rsGrnInv = Nothing
	%> 
		<tr bgcolor="#F4F4A0">	 
			<% if type_1 = "D" Then%>
				<td align="center" colspan="5"><b>GRAND TOTAL</b></td>
			<%Else%>
				<td align="center" colspan="4"><b>GRAND TOTAL</b></td>
			<%End If%>
			<td align="right"><b><%=formatnumber(nbucket,0)%></td>
			<td align="right"><b><%=formatnumber(nbbr,2)%></td>
			<td align="right"><b><%=formatnumber(ntotalkg,2)%></td>
			<td align="right"><b><%=formatnumber(namount,2)%></td> 
		</tr>
	</tbody>
</table>
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