<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
	
	ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

 	v_txtFcba = Request("txtFcba") 
 	v_txtACCOUNTINGYEARPERIOD=trim(request("ACCOUNTINGYEARPERIOD"))
	v_txtBANKCODE = trim(request("BANKCODE")) 

	varDefaultBa = Session("fcBA")   
	v_txtdatefrom = Request("txtdatefrom")
	v_txtdateuntil = Request("txtdateuntil")
	If (v_txtdatefrom = "" Or Len(Trim(v_txtdatefrom)) = 0) Then v_txtdatefrom = textdate(date)
	If (v_txtdateuntil = "" Or Len(Trim(v_txtdateuntil)) = 0) Then v_txtdateuntil = textdate(date)

	If "" & txttdate = "" Then txttdate = textDate(date)
	txttdate = txtdatefrom & "" & txtdateuntil

	If (v_txtFcba = "" Or Len(Trim(v_txtFcba)) = 0) Then
		v_txtFcba = varDefaultBa
	End If	

	q=	"SELECT CONTROLJOB FROM BANK WHERE FCCODE ='" & v_txtBANKCODE & _
	  	"' AND  FCBA= '" & Session("fcBA") & "'"
	Set rsCARI = Server.CreateObject("ADODB.Recordset")
	rsCARI.CursorType = 3
	rsCARI.CursorLocation = 3
	rsCARI.Open q, connstr	 
	if Not rsCARI.eof Then
	   txtControlJob=rsCARI("CONTROLJOB")
	else
	   txtControlJob=""   
	End if
	rsCari.Close
	set rsCari=nothing
		
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
	
		strGrnInv	=	" SELECT fcba," & _
						"        datecreated," & _
						"        bankcode," & _
						"        controljob," & _
						"	     job_api.get_jobcode_name(jobcode, fcba) jobname," & _
						"        transtype," & _
						"        vouchercode," & _
						"        orang," & _
						"        jobcode," & _
						"        locationtype," & _
						"        locationcode," & _
						"        nvl(saldo_awal_per_hari,0) saldo_awal_per_hari," & _
						"        debet," & _
						"        credit," & _
						"        REFERENCE," & _
						"        remarks" & _
						"   FROM (SELECT pv.fcba," & _
						"                pv.datecreated," & _
						"                pv.bankcode," & _
						"                pv.controljob," & _
						"                pv.transtype," & _
						"                pv.vouchercode," & _
						"                pv.orang," & _
						"                pv.jobcode," & _
						"                pv.locationtype," & _
						"                pv.locationcode," & _
						"                pv.lineno," & _
						"                tb_monthly.OPENING_BALANCE_Monthly," & _
						"                NVL (total_pv_daily.amount_yesterday_pv, 0)" & _
						"                   AS amount_yesterday_pv," & _
						"                CASE" & _
						"                   WHEN NVL (total_pv_daily.amount_yesterday_pv, 0) = 0" & _
						"                   THEN" & _
						"                      tb_monthly.OPENING_BALANCE_Monthly" & _
						"                   ELSE" & _
						"                      NVL (" & _
						"                         (  tb_monthly.OPENING_BALANCE_Monthly" & _
						"                          + total_pv_daily.amount_yesterday_pv)," & _
						"                         0)" & _
						"                END" & _
						"                   AS saldo_awal_per_hari," & _
						"                pv.debet," & _
						"                pv.credit," & _
						"                pv.REFERENCE," & _
						"                pv.remarks" & _
						"           FROM (  SELECT iplasprod.paymentvoucher.fcba," & _
						"                          iplasprod.paymentvoucher.datecreated," & _
						"                          iplasprod.paymentvoucher.bankcode," & _
						"                          iplasprod.bank.controljob," & _
						"                          iplasprod.paymentvoucher.transtype," & _
						"                          iplasprod.paymentvoucherdetail.vouchercode," & _
						"                          iplasprod.paymentvoucher.remarks orang," & _
						"                          iplasprod.paymentvoucherdetail.jobcode," & _
						"                          iplasprod.paymentvoucherdetail.locationtype," & _
						"                          iplasprod.paymentvoucherdetail.locationcode," & _
						"                          iplasprod.paymentvoucherdetail.lineno," & _
						"                          NVL (SUM (iplasprod.paymentvoucherdetail.amount), 0)" & _
						"                             AS debet," & _
						"                          0 AS credit," & _
						"                          iplasprod.paymentvoucherdetail.REFERENCE," & _
						"                          iplasprod.paymentvoucherdetail.remarks" & _
						"                     FROM iplasprod.paymentvoucherdetail" & _
						"                          INNER JOIN iplasprod.paymentvoucher" & _
						"                             ON (    iplasprod.paymentvoucherdetail.fcba =" & _
						"                                        iplasprod.paymentvoucher.fcba" & _
						"                                 AND iplasprod.paymentvoucherdetail.vouchercode =" & _
						"                                        iplasprod.paymentvoucher.vouchercode)" & _
						"                          INNER JOIN iplasprod.bank" & _
						"                             ON (    iplasprod.bank.fcba =" & _
						"                                        iplasprod.paymentvoucher.fcba" & _
						"                                 AND iplasprod.bank.fccode =" & _
						"                                        iplasprod.paymentvoucher.bankcode)" & _
						"                    WHERE     iplasprod.paymentvoucher.transtype = '01-RECEIVE'" & _
						"                          AND iplasprod.paymentvoucher.rowstateid <> 'Rejected'" & _
						"		   				   AND iplasprod.paymentvoucher.datecreated >= to_Char(TO_DATE('" & v_txtdatefrom & "','fmMM/dd/yyyy'))" & _
						" 		   				   AND iplasprod.paymentvoucher.datecreated <= to_Char(TO_DATE('" & v_txtdateuntil & "','fmMM/dd/yyyy'))" & _
						"                 GROUP BY iplasprod.paymentvoucher.fcba," & _
						"                          iplasprod.paymentvoucher.datecreated," & _
						"                          iplasprod.paymentvoucher.bankcode," & _
						"                          iplasprod.bank.controljob," & _
						"                          iplasprod.paymentvoucher.transtype," & _
						"                          iplasprod.paymentvoucherdetail.vouchercode," & _
						"                          iplasprod.paymentvoucher.remarks," & _
						"                          iplasprod.paymentvoucherdetail.jobcode," & _
						"                          iplasprod.paymentvoucherdetail.locationtype," & _
						"                          iplasprod.paymentvoucherdetail.locationcode," & _
						"                          iplasprod.paymentvoucherdetail.lineno," & _
						"                          iplasprod.paymentvoucherdetail.REFERENCE," & _
						"                          iplasprod.paymentvoucherdetail.remarks" & _
						"                 UNION ALL" & _
						"                   SELECT iplasprod.paymentvoucher.fcba," & _
						"                          iplasprod.paymentvoucher.datecreated," & _
						"                          iplasprod.paymentvoucher.bankcode," & _
						"                          iplasprod.bank.controljob," & _
						"                          iplasprod.paymentvoucher.transtype," & _
						"                          iplasprod.paymentvoucherdetail.vouchercode," & _
						"                          iplasprod.paymentvoucher.remarks orang," & _
						"                          iplasprod.paymentvoucherdetail.jobcode," & _
						"                          iplasprod.paymentvoucherdetail.locationtype," & _
						"                          iplasprod.paymentvoucherdetail.locationcode," & _
						"                          iplasprod.paymentvoucherdetail.lineno," & _
						"                          0 AS debet," & _
						"                          NVL (SUM (iplasprod.paymentvoucherdetail.amount), 0)" & _
						"                             AS credit," & _
						"                          iplasprod.paymentvoucherdetail.REFERENCE," & _
						"                          iplasprod.paymentvoucherdetail.remarks" & _
						"                     FROM iplasprod.paymentvoucherdetail" & _
						"                          INNER JOIN iplasprod.paymentvoucher" & _
						"                             ON (    iplasprod.paymentvoucherdetail.fcba =" & _
						"                                        iplasprod.paymentvoucher.fcba" & _
						"                                 AND iplasprod.paymentvoucherdetail.vouchercode =" & _
						"                                        iplasprod.paymentvoucher.vouchercode)" & _
						"                          INNER JOIN iplasprod.bank" & _
						"                             ON (    iplasprod.bank.fcba =" & _
						"                                        iplasprod.paymentvoucher.fcba" & _
						"                                 AND iplasprod.bank.fccode =" & _
						"                                        iplasprod.paymentvoucher.bankcode)" & _
						"                    WHERE     iplasprod.paymentvoucher.transtype = '02-PAYMENT'" & _
						"                          AND iplasprod.paymentvoucher.rowstateid <> 'Rejected'" & _
						"		   				   AND iplasprod.paymentvoucher.datecreated >= to_Char(TO_DATE('" & v_txtdatefrom & "','fmMM/dd/yyyy'))" & _
						" 		   				   AND iplasprod.paymentvoucher.datecreated <= to_Char(TO_DATE('" & v_txtdateuntil & "','fmMM/dd/yyyy'))" & _
						"                 GROUP BY iplasprod.paymentvoucher.fcba," & _
						"                          iplasprod.paymentvoucher.datecreated," & _
						"                          iplasprod.paymentvoucher.bankcode," & _
						"                          iplasprod.bank.controljob," & _
						"                          iplasprod.paymentvoucher.transtype," & _
						"                          iplasprod.paymentvoucherdetail.vouchercode," & _
						"                          iplasprod.paymentvoucher.remarks," & _
						"                          iplasprod.paymentvoucherdetail.jobcode," & _
						"                          iplasprod.paymentvoucherdetail.locationtype," & _
						"                          iplasprod.paymentvoucherdetail.locationcode," & _
						"                          iplasprod.paymentvoucherdetail.lineno," & _
						"                          iplasprod.paymentvoucherdetail.REFERENCE," & _
						"                          iplasprod.paymentvoucherdetail.remarks) pv" & _
						"                LEFT JOIN" & _
						"                (  SELECT fcba," & _
						"                          period_year," & _
						"                          period_month," & _
						"                          fccode," & _
						"                          NVL (SUM (OPENING_BALANCE), 0)" & _
						"                             AS OPENING_BALANCE_Monthly" & _
						"                     FROM TRIALBALANCE" & _
						"                 GROUP BY fcba," & _
						"                          fccode," & _
						"                          period_year," & _
						"                          period_month," & _
						"                          fccode) tb_monthly" & _
						"                   ON (    tb_monthly.fcba = pv.fcba" & _
						"                       AND tb_monthly.fccode = pv.controljob" & _
						"                       AND tb_monthly.period_month =" & _
						"                              TO_NUMBER (TO_CHAR (pv.datecreated, 'mm'))" & _
						"                       AND tb_monthly.period_year =" & _
						"                              TO_NUMBER (TO_CHAR (pv.datecreated, 'yyyy')))" & _
						"                LEFT JOIN" & _
						"                (  SELECT fcba," & _
						"                          bankcode," & _
						"                          controljob," & _
						"                          SUM (amount_yesterday_pv) AS amount_yesterday_pv" & _
						"                     FROM (  SELECT iplasprod.paymentvoucher.fcba," & _
						"                                    iplasprod.paymentvoucher.datecreated," & _
						"                                    iplasprod.paymentvoucher.bankcode," & _
						"                                    iplasprod.bank.controljob," & _
						"                                    iplasprod.paymentvoucher.transtype," & _
						"                                    CASE" & _
						"                                       WHEN iplasprod.paymentvoucher.transtype =" & _
						"                                               '02-PAYMENT'" & _
						"                                       THEN" & _
						"                                            NVL (" & _
						"                                               SUM (" & _
						"                                                  iplasprod.paymentvoucherdetail.amount)," & _
						"                                               0)" & _
						"                                          * -1" & _
						"                                       ELSE" & _
						"                                          NVL (" & _
						"                                             SUM (" & _
						"                                                iplasprod.paymentvoucherdetail.amount)," & _
						"                                             0)" & _
						"                                    END" & _
						"                                       AS amount_yesterday_pv" & _
						"                               FROM iplasprod.paymentvoucherdetail" & _
						"                                    INNER JOIN iplasprod.paymentvoucher" & _
						"                                       ON (    iplasprod.paymentvoucherdetail.fcba =" & _
						"                                                  iplasprod.paymentvoucher.fcba" & _
						"                                           AND iplasprod.paymentvoucherdetail.vouchercode =" & _
						"                                                  iplasprod.paymentvoucher.vouchercode)" & _
						"                                    INNER JOIN iplasprod.bank" & _
						"                                       ON (    iplasprod.bank.fcba =" & _
						"                                                  iplasprod.paymentvoucher.fcba" & _
						"                                           AND iplasprod.bank.fccode =" & _
						"                                                  iplasprod.paymentvoucher.bankcode)" & _
						"                              WHERE     iplasprod.paymentvoucher.transtype IN ('01-RECEIVE', '02-PAYMENT')" & _
						"                           AND iplasprod.paymentvoucher.rowstateid ='Rejected'" & _
						"		   				  	AND iplasprod.paymentvoucher.datecreated >= to_Char(TO_DATE('" & v_txtdatefrom & "','fmMM/dd/yyyy'))" & _
						" 		   				   	AND iplasprod.paymentvoucher.datecreated <= to_Char(TO_DATE('" & v_txtdateuntil & "','fmMM/dd/yyyy'))" & _
						"                           GROUP BY iplasprod.paymentvoucher.fcba," & _
						"                                    iplasprod.paymentvoucher.datecreated," & _
						"                                    iplasprod.paymentvoucher.bankcode," & _
						"                                    iplasprod.bank.controljob," & _
						"                                    iplasprod.paymentvoucher.transtype)" & _
						"                    WHERE " & _    
						"		   				  	datecreated >= to_Char(TO_DATE('" & v_txtdatefrom & "','fmMM/dd/yyyy'))" & _
						" 		   				   	AND datecreated <= to_Char(TO_DATE('" & v_txtdateuntil & "','fmMM/dd/yyyy'))" & _
						"                 GROUP BY fcba, bankcode, controljob) total_pv_daily" & _
						"                   ON (    total_pv_daily.fcba = pv.fcba" & _
						"                       AND total_pv_daily.bankcode = pv.bankcode" & _
						"                       AND total_pv_daily.controljob = pv.controljob)" & _
						"          WHERE     pv.fcba = '"& session("fcba") &"'" & _
						"		   AND  pv.datecreated >= to_Char(TO_DATE('" & v_txtdatefrom & "','fmMM/dd/yyyy'))" & _
						" 		   AND  pv.datecreated <= to_Char(TO_DATE('" & v_txtdateuntil & "','fmMM/dd/yyyy'))" & _
						" 		   AND pv.bankcode ='"& v_txtBANKCODE &"')"

			 
			if  len(trim(v_txtdatefrom)) > 0 then 
				txtSelect =  " From " & FormatDateTime(v_txtdatefrom,1)
			end if
			if  len(trim(v_txtdateuntil)) > 0 then
				txtSelect =  txtSelect & " to " & FormatDateTime(v_txtdateuntil,1)
			end if
			if  len(trim(v_txtBANKCODE)) > 0 then
				txtSelect =  txtSelect & " Bank Code " & v_txtBANKCODE
			end if
			if len(trim(txtSelect)) > 0 then
			   txtSelect = "Selected  by : " & txtSelect
			else
			   txtSelect = "Selected All Data"
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
								
	strGrnInv	=	strGrnInv & " ORDER BY fcba," & _
								"        datecreated," & _
								"        transtype," & _
								"        vouchercode," & _
								"		 lineno" 

	'response.write strGrnInv	
	Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
	rsGrnInv.Open strGrnInv, connstr, adOpenStatic 
	
%>


<html>

<head>
<title>ePlaS Reporting - Arus Cash Harian</title>
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
$( ".datepicker" ).datepicker();
});
</script>
<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}
 

</script> 

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
							<font face="verdana" size="1">From</font>
						</td>
						<td width="4">
							<input class="datepicker" type="text" maxlength="10" name="txtdatefrom" size="10" value="<%=v_txtdatefrom%>">
						</td>
						<td>
							<font face="verdana" size="1">To</font>
						</td>
						<td width="2">
							<input class="datepicker" type="text" maxlength="10" name="txtdateuntil" size="10" value="<%=v_txtdateuntil%>">
						</td>  
				     	<td>
							<font face="Verdana" size="1">Bank</font>
							</td>
							<td>
							   	<select name="BANKCODE">
							   		<option></option>
									<%
									q = " SELECT a.fccode, a.fcname" & _
										"  FROM bank a" & _
										" WHERE a.fcba = '"& session ("fcBA") &"'" & _
										"       AND a.activation = 'Y'" & _
									    "       AND UPPER(a.banktype) <> ('JOURNALRECLASS')" 

									if UCase(cstr(""&banktype)) = "JOURNAL" then
										q = q & "       AND UPPER(banktype) = UPPER('JOURNAL')"
									else
										q = q & "       AND UPPER(banktype) IN ('CASH', 'BANK')"
									end if
									' response.write q
									Set rsComboA = Server.CreateObject("ADODB.Recordset")
									
									rsComboA.Open q, connstr
									if not rsComboA.eof then
										Do While Not rsComboA.eof
									%>
										<option value="<%=rsComboA("fccode")%>" <% If cstr(""&v_txtBANKCODE) = cstr(""&rsComboA("fccode")) Then %> Selected <%End if%>><%=rsComboA("fcname")%></option> 
									<% 
											rsComboA.movenext
										loop
									End if
									rsComboA.Close : set rsComboA = Nothing : q = ""
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
	    	<th colspan="8" style="background-color:#FFFFFF; border:none;">  
				<div id="header-wrapper">
					<div id="head_judul" style="float:left; text-align:left;">
						<h4>PT. <%=CompanyName%></h4>
						<h5><%=fcbaname%></h5>
					</div>
					<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
						<h2>ARUS CASH HARIAN</h2>
		                <br/>
					</div>
					<div align="right">
						<a href="javascript:launchwin('print_aruscash_harian.Asp?ACCOUNTINGYEARPERIOD=<%=v_txtACCOUNTINGYEARPERIOD%>&BANKCODE=<%=v_txtBANKCODE%>&txtdatefrom=<%=v_txtdatefrom%>&txtdateuntil=<%=v_txtdateuntil%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
							<img src="../IMAGES/Print.gif" border="0" title="Print Preview">
							<font face="Verdana" size="1">Print</font>
						</a>&nbsp;&nbsp;
					</div>
				</div>
			</th>
		</tr>
		<tr>
			<th colspan="11" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
		</tr>
		<tr>
			<th> Date</th>
			<th> Voucher Code </th>
			<th> Paid For / Receive From</th>  
			<th> Reference</th>
			<th> Remarks</th>
			<th> Debet</th>
			<th> Credit</th>
			<th> Saldo</th>
		</tr>
	</thead>
<tbody>
<%
	Dim v_Opening,v_Debit,v_Credit,v_Saldo	
	if Not rsGrnInv.eof then  
	
	   nCountofField=rsGrnInv.fields.count-1
	v_Amount = 0
	v_debet = 0 
	v_credit = 0
	v_Saldo = 0
%>
		<tr>
			<td align="left" style="<%=cmdstyleLeft%>" colspan="7"><font face="Verdana" size="1">Opening Balance.........</font></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsGrnInv("saldo_awal_per_hari"))),2)%></font></td>
		</tr>
	
	<%	v_Amount=cdbl(to_nol(rsGrnInv("saldo_awal_per_hari")))
		v_Saldo=cdbl(to_nol(rsGrnInv("saldo_awal_per_hari")))
		while not rsGrnInv.eof
			v_Saldo= v_Saldo + cdbl(to_nol(rsGrnInv("debet"))) - cdbl(to_nol(rsGrnInv("credit")))
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
	%> 
		<tr class="<%=cClass%>">
			<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsGrnInv("Datecreated")%></td>
			<td align="Center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsGrnInv("Vouchercode")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsGrnInv("orang")%></td> 
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsGrnInv("Reference")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rsGrnInv("Remarks")%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsGrnInv("debet"))),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(rsGrnInv("credit"))),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_Saldo)),2)%></td>
		</tr>
<%
			v_Amount= v_Amount + cdbl(to_nol(rsGrnInv("debet"))) - cdbl(to_nol(rsGrnInv("credit")))
			v_debet = v_debet +  cdbl(to_nol(rsGrnInv("debet"))) 
			v_credit = v_credit +  cdbl(to_nol(rsGrnInv("credit"))) 
			rsGrnInv.MoveNext
		Wend
	rsGrnInv.Close
	set rsGrnInv = Nothing
%>
	</tbody>
			<input type="hidden" name="ACCOUNTINGYEARPERIOD" value="<%=v_txtACCOUNTINGYEARPERIOD%>">
			<input type="hidden" name="BANKCODE" value="<%=v_txtBANKCODE%>">
		<tr>
			<td align="left" style="<%=cmdstyleLeft%>" colspan="5"> <font face="Verdana" size="1">Ending Balance.........</td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_debet)),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_credit)),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_Amount)),2)%></td>
		</tr>
<% else %>
		<tr>
			<td colspan="8" align="left">NO DATA AVAILABLE</td>
		</tr>
<% end if %>
</table>
<table width="100%" class="data" align="center">
	<tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td> 
	</tr>
	<tr>
		<td align="center"><font face="Verdana" size="1" style="font-weight: bold;">PREPARED BY,</td>
		<td align="center"><font face="Verdana" size="1" style="font-weight: bold;">CHECKED BY,</td>
		<td align="center"><font face="Verdana" size="1" style="font-weight: bold;">APPROVED BY,</td>
	</tr>
	<tr> 
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td> 
	</tr>
	<tr> 
		<td align="center"><font face="Verdana" size="1" style="font-weight: bold;"><br/><br/>KASIR</td>
		<td align="center"><font face="Verdana" size="1" style="font-weight: bold;"><br/><br/><%=Session("KTU")%></td>
		<td align="center"><font face="Verdana" size="1" style="font-weight: bold;"><br/><br/><%=Session("FCMANAGER")%></td>
	</tr>
</table>
</body>
</html>