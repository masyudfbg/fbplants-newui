<%validatesession%>
<!--#include file="../connDb1.Inc"-->
<!--#include file="../txtFunction.inc"-->  
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ePlaS - Payment Voucher</title>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script type="text/javascript" language="JavaScript1.2" src="stlib.js"></script>
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
		
	<script language="JavaScript" src="/js/jquery.printPage.js"></script> 
	<script>
					function toggle(id){
						var e = document.getElementById(id);
						if(e.style.display == 'inline-block'){
							e.style.display = 'none';
						}else if(e.style.display == 'none'){
							e.style.display = 'inline-block';
						}
					}
	</script>
	<script>
	 function printDiv(divName) {
	 var printContents = document.getElementById(divName).innerHTML;
     var originalContents = document.body.innerHTML;

     document.body.innerHTML = printContents;

     window.print();

     document.body.innerHTML = originalContents;
}
</script>
<style type="text/css">
@media all {
#product-table td{padding:5px;}
.page-break	{ display: block; }
a.button_print { display:block; }	
#wrapper{background: #e2e2e2; /* Old browsers */
background: -moz-linear-gradient(-45deg,  #e2e2e2 0%, #dbdbdb 50%, #d1d1d1 51%, #fefefe 100%); /* FF3.6+ */
background: -webkit-gradient(linear, left top, right bottom, color-stop(0%,#e2e2e2), color-stop(50%,#dbdbdb), color-stop(51%,#d1d1d1), color-stop(100%,#fefefe)); /* Chrome,Safari4+ */
background: -webkit-linear-gradient(-45deg,  #e2e2e2 0%,#dbdbdb 50%,#d1d1d1 51%,#fefefe 100%); /* Chrome10+,Safari5.1+ */
background: -o-linear-gradient(-45deg,  #e2e2e2 0%,#dbdbdb 50%,#d1d1d1 51%,#fefefe 100%); /* Opera 11.10+ */
background: -ms-linear-gradient(-45deg,  #e2e2e2 0%,#dbdbdb 50%,#d1d1d1 51%,#fefefe 100%); /* IE10+ */
background: linear-gradient(135deg,  #e2e2e2 0%,#dbdbdb 50%,#d1d1d1 51%,#fefefe 100%); /* W3C */
filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#e2e2e2', endColorstr='#fefefe',GradientType=1 ); /* IE6-9 fallback on horizontal gradient */
}
a.button_print {display: inline-block;}
.letter { 
background-color:#fff; 
box-shadow:0px 0px 4px rgba(0,0,0,0.4); 
border-radius:4px;
margin: auto;
width: 1200px;
padding-top:30px;
padding-bottom:30px;
padding-left:30px;
padding-right:30px;
}
#header-wrapper {display:block} 
} 
</style>		
</head>

<body>
<%
cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

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
	
	q=	"SELECT BANKTYPE FROM BANK WHERE FCCODE ='" & v_txtBANKCODE & _
	  	"' AND  FCBA= '" & Session("fcBA") & "'"
	Set rsCARI = Server.CreateObject("ADODB.Recordset")
	rsCARI.CursorType = 3
	rsCARI.CursorLocation = 3
	rsCARI.Open q, connstr	 
	if Not rsCARI.eof Then
	   BANKTYPE=rsCARI("BANKTYPE")
	else
	   BANKTYPE=""   
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
						"        saldo_awal_per_hari," & _
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
						"                          AND iplasprod.paymentvoucher.rowstateid = 'Approved'" & _
						"                          AND    iplasprod.paymentvoucher.accountingyear" & _
						"                              || LPAD (" & _
						"                                    iplasprod.paymentvoucher.accountingperiod," & _
						"                                    2," & _
						"                                    0) = '"& v_txtACCOUNTINGYEARPERIOD &"'" & _
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
						"                          AND iplasprod.paymentvoucher.rowstateid = 'Approved'" & _
						"                          AND    iplasprod.paymentvoucher.accountingyear" & _
						"                              || LPAD (" & _
						"                                    iplasprod.paymentvoucher.accountingperiod," & _
						"                                    2," & _
						"                                    0) = '"& v_txtACCOUNTINGYEARPERIOD &"'" & _
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
						"                INNER JOIN" & _
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
						"                              WHERE     iplasprod.paymentvoucher.transtype IN" & _
						"                                           ('01-RECEIVE', '02-PAYMENT')" & _
						"                                    AND iplasprod.paymentvoucher.rowstateid =" & _
						"                                           'Approved'" & _
						"                                    AND    iplasprod.paymentvoucher.accountingyear" & _
						"                                        || LPAD (" & _
						"                                              iplasprod.paymentvoucher.accountingperiod," & _
						"                                              2," & _
						"                                              0) = '"& v_txtACCOUNTINGYEARPERIOD &"'" & _
						"                           GROUP BY iplasprod.paymentvoucher.fcba," & _
						"                                    iplasprod.paymentvoucher.datecreated," & _
						"                                    iplasprod.paymentvoucher.bankcode," & _
						"                                    iplasprod.bank.controljob," & _
						"                                    iplasprod.paymentvoucher.transtype)" & _
						"                    WHERE 		TO_CHAR (datecreated, 'yyyymm') = '"& v_txtACCOUNTINGYEARPERIOD &"'" & _    
						"						   AND TO_NUMBER (TO_CHAR (datecreated, 'dd')) >= 01" & _
						"                          AND TO_NUMBER (TO_CHAR (datecreated, 'dd')) < "& day(v_txtdatefrom) &"" & _
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
						if  len(trim(v_txtACCOUNTINGYEARPERIOD)) > 0 then
							txtSelect = txtSelect &  " ACCOUNTINGYEARPERIOD " &  v_txtACCOUNTINGYEARPERIOD
						end if
						if  len(trim(v_txtBANKCODE)) > 0 then
							txtSelect =  txtSelect & " Bank Code " & v_txtBANKCODE
						end if
						if len(trim(txtSelect)) > 0 then
						   txtSelect = "Selected  by : " & txtSelect
						else
						   txtSelect = "Selected All Data"
						end if
			

	 

	' strGrnInv	=	strGrnInv &	" GROUP BY" & _
	' 							"     FCBA" & _
	' 							"     , NOPO" & _
	' 							"     , STOCKCODE"
								
	strGrnInv	=	strGrnInv & " ORDER BY fcba," & _
								"        datecreated," & _
								"        transtype," & _
								"        vouchercode," & _
								"		 lineno" 

	' response.write strGrnInv	
	Set rsGrnInv = Server.CreateObject("ADODB.Recordset")
	rsGrnInv.Open strGrnInv, connstr, adOpenStatic 

%>
<div id="wrapper">
<div class="navi_to_print">
<input type="button" value="Print" onclick="printDiv('note')" />

</div>
<div class="letter">
<div id="note">
<style>
@media all {
td{padding:5px;}
table.PAYSLIP td{padding:5px;}
.page-break hr {display:block;border: 1px dashed red;}
.PAYSLIP td {font-family:verdana,sans-serif; font-size:8px; padding:0px;}
#note {margin-left:auto;padding:5px;}
#info_note .judul { text-align:center; font-size:16px;padding-top:0px;padding-bottom:0px;}
#info_note .left { float: left; width:35%;}
#info_note .right {  float: right; width:35%;}
#info_note .point { font-size:10px; font-weight:bold; width:100%;  margin-top:5px;}
#info_note .point span { font-size:10px; font-style:italic; display:block; border-top:1px #000 solid;}
#info_note .point .name { width:100px; float:left; margin-top :10px;}
#info_note .point .isi  { float:left; height:26px; line-height:25px;margin-top:7px; margin-left:10px;font-size:10px;}

#info_note .point_right { font-size:10px; font-weight:bold; width:100%;  margin-top:5px; text-align:left;}
#info_note .point_right span { font-size:10px; font-style:italic; display:block; border-top:1px #000 solid;}
#info_note .point_right .name {width:25%; float:left;  margin-top :10px;}
#info_note .point_right .isi  { float:left; height:26px; line-height:25px; margin-left:10px;margin-top:7px;font-size:10px;}
.clear { clear:both; content:'';}
table .long {width:400px;}
table .small {width:40px;}
table th span { font-style:italic; font-weight:normal;}
.line { height:1px; background:#000; content:' ';margin: auto;}
.line2 { width:90%; height:1px; background:#000000; content:' '; margin:auto;}
#product-table th { background-color:#fff; color:#000;}
#product-table td { border-color:#000; color:#000; font-size:9px;}
#product-table td.none { border: none;}
#footer_note .block { font-size:10px; float:left; width:25%; height:90px; background-color:#fff; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:90%; text-align:center;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
.terbilang{font-size:10px;float:left; text-align:left;font-style: italic;margin-left:20px; padding-top:10px;}
.title_terbilang {width:100px;float:left; text-align:left;font-style: italic;font-size:10px;}
.page{width:100%;}
}
@media print
{
td{padding:5px;}
table.PAYSLIP td{padding:5px;}
.page-break hr {display:block;border: 1px dashed red;}
.PAYSLIP td {font-family:verdana,sans-serif; font-size:8px; padding:0px;}
#note {margin-left:auto;padding:5px;}
#info_note .judul { text-align:center; font-size:16px;padding-top:0px;padding-bottom:0px;}
#info_note .left { float: left; width:35%;}
#info_note .right {  float: right; width:35%;}
#info_note .point { font-size:10px; font-weight:bold; width:100%;  margin-top:5px;}
#info_note .point span { font-size:10px; font-style:italic; display:block; border-top:1px #000 solid;}
#info_note .point .name { width:100px; float:left; margin-top :10px;}
#info_note .point .isi  { float:left; height:26px; line-height:25px;margin-top:7px; margin-left:10px;font-size:10px;}

#info_note .point_right { font-size:10px; font-weight:bold; width:100%;  margin-top:5px; text-align:left;}
#info_note .point_right span { font-size:10px; font-style:italic; display:block; border-top:1px #000 solid;}
#info_note .point_right .name {width:25%; float:left;  margin-top :10px;}
#info_note .point_right .isi  { float:left; height:26px; line-height:25px; margin-left:10px;margin-top:7px;font-size:10px;}
.clear { clear:both; content:'';}
table .long {width:400px;}
table .small {width:40px;}
table th span { font-style:italic; font-weight:normal;}
.line { height:1px; background:#000; content:' ';margin: auto;}
.line2 { width:90%; height:1px; background:#000000; content:' '; margin:auto;}
#product-table th { background-color:#fff; color:#000;}
#product-table td { border-color:#000; color:#000; font-size:9px;}
#product-table td.none { border: none;}
#footer_note .block { font-size:10px; float:left; width:19%; height:90px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
.terbilang{font-size:10px;float:left; text-align:left;font-style: italic;  padding-top:10px;}
.title_terbilang {width:100px;float:left; text-align:left;font-style: italic;font-size:10px;}
.page{width:98.5%;}
}

</style>
<div class="page" >  
	<table border="0" width="100%" id="product-table">
		<thead>
			<tr>
				<th colspan="11" style=" background-color:#FFFFFF; color:#000000; border:none">
					<div id="header-wrapper">
					   <div id="head_judul" style="float:left; text-align:left;">
							<h4>PT. <%=Session("COMPANY")%></h4>
							<h5><%=Session("BA")%></h5>
						</div>
						<div id="subheader" class="centered" style="width:50%; margin:auto;">
							<%If cstr(""&BANKTYPE) = "BANK" Then %>
								<h1 class="judul"> BANK BOOK </h1>
							<%Else%>
								<h1 class="judul"> CASH BOOK </h1>
							<%End If%>
						</div> 	
						<div style="clear:both;"></div>
					</div>
				</th>
			</tr> 
			<!--<div class="printhide">--> 
			<tr>
				<th colspan="7" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=txtSelect%></th>
			</tr>
			<!--</div>--> 
				<th> Date</th>
				<th width="200"> Voucher Code </th>
				<th> Paid For / Receive From</th>  
				<th> Remarks</th>
				<th> Debet</th>
				<th> Credit</th>
				<th> Saldo</th>
			</tr>
		</thead>
	<tbody>
<%
	Dim v_Opening,v_Debit,v_Credit,v_Saldo	
	if  len(trim(v_txtACCOUNTINGYEARPERIOD)) > 0 then
        Set rsGrnInv = connStr.Execute(strGrnInv)
	
	
	if Not rsGrnInv.eof then
	   nCountofField=rsGrnInv.fields.count-1
	end if
	v_Amount = 0
	v_debet = 0 
	v_credit = 0
	v_Saldo = 0
%>
		<tr>
			<td align="left" style="<%=cmdstyleLeft%>" colspan="6"><font face="Verdana" size="1">Opening Balance.........</font></td>
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
			<td align="left" style="<%=cmdstyleLeft%>" colspan="4"> <font face="Verdana" size="1">Ending Balance.........</td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_debet)),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_credit)),2)%></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><%=FormatNumber(cdbl(to_nol(v_Amount)),2)%></td>
		</tr>
<% else %>
		<tr>
			<td colspan="21" align="left">NO DATA AVAILABLE</td>
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






