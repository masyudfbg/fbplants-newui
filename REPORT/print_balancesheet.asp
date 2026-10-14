<%validatesession%>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ePlaS Reporting - Balance Sheet</title>
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
/*width: 1200px;*/
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
v_txtYear		=	Request("period_year")
v_txtMonth		=	Request("period_month")
v_txtdate		=	Request("v_txtdate")
v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")
v_txtOwnership			= Request("txtOwnership")

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

' if request("txtYear")="" then
'    v_txtYear	=	Year(Session("v_Costbook"))
' end if
' if request("txtMonth")="" then
'    v_txtMonth	=	Month(Session("v_Costbook"))
' end if

' v_txtdate = v_txtYear &""& v_txtMonth
 

' strTipeBa		=	" SELECT FCTYPE FROM BUSINESSUNIT WHERE FCCODE = '" & Session("FCBA") & "'"

' Set rsTipeBa	=	Server.CreateObject("ADODB.Recordset")
' rsTipeBa.Open strTipeBa, connstr, adLockOptimistic, adOpenStatic

' if not rsTipeBa.eof then
' validasiTipeBa	=	rsTipeBa("FCTYPE")
' else
' Response.Write  "<font face=""Verdana"" size=""1"" color=""red"">" & "iPla'S Error: Business Area ( Type ) is Empty" & "</font>"
' err.number		=	1101
' validasiTipeBa	=	""
' end if
' rsTipeBa.close
' set rsTipeBa = Nothing
' set strTipeBa = Nothing
strDefaultBa = 	" SELECT" & _
				"     FCCODE as FCBA" & _
				" FROM BUSINESSUNIT BU" & _
				" INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) " & _
				" WHERE" & _
				"     FCTYPE = 'HO'" & _
				"     AND ISACTIVE= 'TRUE'" & _
				"     AND FCUSERCODE = '" & Session("CodeName") & "'"
' response.write strDefaultBa
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

strFieldCons	=	" USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',STAGE1.fcba) GroupPlantation" & _
					" , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',STAGE1.fcba) Plantation" & _
					" , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',STAGE1.fcba) groupregion" & _
					" , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',STAGE1.fcba) region" & _
					" , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',STAGE1.fcba) company" & _
					" , BUSINESSUNIT_API.get_OrderData(STAGE1.fcba) NoUrut, STAGE1.fcba,"
					
wherecons	=	" where company is not null "

If Len(Trim(v_txtGroupPlantation)) > 0 Then
	if v_txtGroupPlantation = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtGroupPlantation = "--ALL--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " And GroupPlantation = '" & v_txtGroupPlantation & "'"
		orderCons	=	orderCons	& " "
		var_colspan	=	var_colspan	+ 1
	End If
End If

If Len(Trim(v_txtPlantation)) > 0 Then
	if v_txtPlantation = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtPlantation = "--ALL--" then
		selectCons	=	selectCons	& " , Plantation"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , Plantation"
		whereCons	=	whereCons	& " And Plantation = '" & v_txtPlantation & "'"
		orderCons	=	orderCons	& " , PLANTATION_API.get_OrderData(Plantation)"
		var_colspan	=	var_colspan	+ 1
	End If
End If


If Len(Trim(v_txtGroupRegion)) > 0 Then
	if v_txtGroupRegion = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtGroupRegion = "--ALL--" then
		selectCons	=	selectCons	& " , GroupRegion"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , GroupRegion"
		whereCons	=	whereCons	& " And GroupRegion = '" & v_txtGroupRegion & "'"
		orderCons	=	orderCons	& " , GROUPREGION_API.get_OrderData(GroupRegion)"
		var_colspan	=	var_colspan	+ 1
	End If
End If

If Len(Trim(v_txtRegion)) > 0 Then
	if v_txtRegion = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtRegion = "--ALL--" then
		selectCons	=	selectCons	& " , Region"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , Region"
		whereCons	=	whereCons	& " And Region = '" & v_txtRegion & "'"
		orderCons	=	orderCons	& " , REGION_API.get_OrderData(Region)"
		var_colspan	=	var_colspan	+ 1
	End If
End If


If Len(Trim(v_txtCompany)) > 0 Then
	if v_txtCompany = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtCompany = "--ALL--" then
		selectCons	=	selectCons	& " , company"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
	Else
		selectCons	=	selectCons	& " , company"
		whereCons	=	whereCons	& " And company = '" & v_txtCompany & "'"
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		
		q	=	" SELECT FCCODE" & _
				" , FCCOMPANYCODE" & _
				" , FCNAME" & _
				" , FCCOMPANYNAME" & _
				" , FCMANAGER" & _
				" , FCKTU" & _
				" , FCCROP" & _
				" , FCTYPE" & _
				" FROM BUSINESSUNIT WHERE" & _
				" FCCOMPANYCODE = '" & v_txtCompany & "'"	
		set rsAktiva = connStr.Execute(q)		
		if not rsAktiva.eof then
			CompanyName = rsAktiva("FCCOMPANYNAME")	
		end if
		q	=	""
		rsAktiva.Close
		set rsAktiva = Nothing
	End If
End If

If Len(Trim(v_txtfcba)) = 0 Then v_txtfcba = "--SUMMARY--"

If Len(Trim(v_txtfcba)) > 0 Then
	if v_txtfcba = "--SUMMARY--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		whereCons2	=	wherecons2	& " "
		whereCons3	=	wherecons3	& " "
		orderCons	=	orderCons	& " "
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtfcba = "--ALL--" then
		selectCons	=	selectCons	& " "
		whereCons	=	wherecons	& " "
		whereCons2	=	wherecons2	& " "
		whereCons3	=	wherecons3	& " "
		orderCons	=	orderCons	& " "
	Else
		selectCons	=	selectCons	& " "
		whereCons	=	whereCons	& " "
		whereCons2	=	wherecons2	& " AND tb.fcba = '"& v_txtfcba &"'"
		whereCons3	=	wherecons3	& " And fcba = '" & v_txtfcba & "'"
		orderCons	=	orderCons	& " , nourut"
		
		q	=	" SELECT FCCODE" & _
				" , FCCOMPANYCODE" & _
				" , FCNAME" & _
				" , FCCOMPANYNAME" & _
				" , FCMANAGER" & _
				" , FCKTU" & _
				" , FCCROP" & _
				" , FCTYPE" & _
				" FROM BUSINESSUNIT WHERE" & _
				" FCCODE = '" & v_txtfcba & "'"	
		set rsAktiva = connStr.Execute(q)		
		if not rsAktiva.eof then
			fcbaname = rsAktiva("FCNAME")
			CompanyName=rsAktiva("FCCOMPANYNAME")
			MgrName = rsAktiva("FCMANAGER")
			KTUName = rsAktiva("FCKTU")
			CropName=rsAktiva("FCCROP")
			FcbaType=rsAktiva("FCTYPE")
		end if
		q	=	""
		rsAktiva.Close
		set rsAktiva = Nothing
	End If
End If


strAktiva	=	"  SELECT "& strFieldCons &"" & _
				"         STAGE1.param_1," & _
				"         STAGE1.param_2," & _
				"         STAGE1.fccode," & _
				"         STAGE1.jobbalancesheetcode," & _
				"         NVL (STAGE2.amount, 0) AKTIVA," & _
				"         NVL (STAGE2.amount_prev, 0) AKTIVA_PREV," & _
				"         NVL (STAGE2.amount_sbi, 0) AKTIVA_SBI," & _
				"         NVL (STAGE2.amount_prev_sbi, 0) AKTIVA_PREV_SBI" & _
				"    FROM (    SELECT fcba," & _
				"                     fccode," & _
				"                     param_1," & _
				"                     param_2," & _
				"                     LPAD('_',5*(level-1),'_') || to_char(fcname) jobbalancesheetcode" & _
				"                FROM parameterdetail" & _
				"               WHERE parheadcode = 'AKTIVA' AND param_1 <> -1  " & _
				"          START WITH fccode = '0'" & _
				"          CONNECT BY PRIOR fccode = param_1" & _
				"            GROUP BY fcba," & _
				"                     fccode," & _
				"                     param_1," & _
				"                     param_2," & _
				"                     LPAD('_',5*(level-1),'_') || to_char(fcname))" & _
				"         STAGE1" & _
				"         LEFT JOIN" & _
				"         (  SELECT tb.fcba, jb.job_forbalancesheetcode, SUM (tb.ending_balance) amount, 0 amount_prev, 0 amount_sbi, 0 amount_prev_sbi" & _
				"              FROM trialbalance tb" & _
				"                   INNER JOIN job jb" & _
				"                      ON (tb.fcba = jb.fcba AND tb.fccode = jb.fccode)" & _
				"             WHERE    period_year = '"& v_txtYear &"'" & _
				"                   AND period_month = '"& v_txtMonth &"'" & _ 
				"          GROUP BY tb.fcba, jb.job_forbalancesheetcode" & _
				" 		   UNION ALL" & _
				"  		   SELECT tb.fcba, jb.job_forbalancesheetcode, 0 amount, SUM (tb.ending_balance) amount_prev, 0 amount_sbi, 0 amount_prev_sbi" & _
				"              FROM trialbalance tb" & _
				"                   INNER JOIN job jb" & _
				"                      ON (tb.fcba = jb.fcba AND tb.fccode = jb.fccode)" & _
				"             WHERE    period_year = '"& v_txtYear &"' - 1" & _
				"                   AND period_month = '"& v_txtMonth &"'" & _ 
				"          GROUP BY tb.fcba, jb.job_forbalancesheetcode" & _
				"		   UNION ALL" & _
				"			  SELECT tb.fcba, jb.job_forbalancesheetcode, 0 amount, 0 amount_prev, SUM (tb.ending_balance) amount_sbi, 0 amount_prev_sbi" & _
				"              FROM trialbalance tb" & _
				"                   INNER JOIN job jb" & _
				"                      ON (tb.fcba = jb.fcba AND tb.fccode = jb.fccode)" & _
				"             WHERE    period_year = '"& v_txtYear &"'" & _
				"                   AND period_month >= '7'" & _ 
				"          GROUP BY tb.fcba, jb.job_forbalancesheetcode" & _
				" 		   UNION ALL" & _
				"  		   SELECT tb.fcba, jb.job_forbalancesheetcode, 0 amount, 0 amount_prev, 0 amount_sbi, SUM (tb.ending_balance) amount_prev_sbi" & _
				"              FROM trialbalance tb" & _
				"                   INNER JOIN job jb" & _
				"                      ON (tb.fcba = jb.fcba AND tb.fccode = jb.fccode)" & _
				"             WHERE    period_year = '"& v_txtYear &"' - 1" & _
				"                   AND period_month >= '7'" & _ 
				"          GROUP BY tb.fcba, jb.job_forbalancesheetcode) STAGE2" & _
				"            ON (    STAGE1.FCBA = STAGE2.FCBA" & _
				"                AND STAGE1.FCCODE = STAGE2.JOB_FORBALANCESHEETCODE)" & _
				" ORDER BY fccode"  

strAktiva2  = 	"          ,param_1," & _
				"          param_2," & _
				"          fccode," & _
				"          jobbalancesheetcode," & _
				"          NVL (SUM (AKTIVA), 0) AKTIVA,"  & _
				"		   NVL (SUM (AKTIVA_PREV), 0) AKTIVA_PREV," & _
				"		   NVL (SUM (AKTIVA_SBI), 0) AKTIVA_SBI," & _
				"		   NVL (SUM (AKTIVA_PREV_SBI), 0) AKTIVA_PREV_SBI"  

strAktivaTotal	=	" select GroupPlantation " & selectCons & strAktiva2 &"  From (" & strAktiva & " )" & _ 
                    " " & whereCons & _
                    " group by GroupPlantation" & selectCons & ",param_1, param_2, fccode, jobbalancesheetcode" & _
                    " ORDER BY fccode"

strPasiva	=	"  SELECT "& strFieldCons &"" & _
				"         STAGE1.param_1," & _
				"         STAGE1.param_2," & _
				"         STAGE1.fccode," & _
				"         STAGE1.jobbalancesheetcode," & _
				"         NVL (STAGE2.amount, 0) PASIVA," & _
				"         NVL (STAGE2.amount_prev, 0) PASIVA_PREV," & _
				"         NVL (STAGE2.amount_sbi, 0) PASIVA_SBI," & _
				"         NVL (STAGE2.amount_prev_sbi, 0) PASIVA_PREV_SBI" & _
				"    FROM (    SELECT fcba," & _
				"                     fccode," & _
				"                     param_1," & _
				"                     param_2," & _
				"                     LPAD('_',5*(level-1),'_') || to_char(fcname) jobbalancesheetcode" & _
				"                FROM parameterdetail" & _
				"               WHERE parheadcode = 'PASIVA'  AND param_1 <> -1  " & _
				"          START WITH fccode = '0'" & _
				"          CONNECT BY PRIOR fccode = param_1" & _
				"            GROUP BY fcba," & _
				"                     fccode," & _
				"                     param_1," & _
				"                     param_2," & _
				"                     LPAD('_',5*(level-1),'_') || to_char(fcname))" & _
				"         STAGE1" & _
				"         LEFT JOIN" & _
				"         (  SELECT tb.fcba, jb.job_forbalancesheetcode, SUM (tb.ending_balance) amount, 0 amount_prev, 0 amount_sbi, 0 amount_prev_sbi" & _
				"              FROM trialbalance tb" & _
				"                   INNER JOIN job jb" & _
				"                      ON (tb.fcba = jb.fcba AND tb.fccode = jb.fccode)" & _
				"             WHERE    period_year = '"& v_txtYear &"'" & _
				"                   AND period_month = '"& v_txtMonth &"'" & _ 
				"          GROUP BY tb.fcba, jb.job_forbalancesheetcode" & _
				" 		   UNION ALL" & _
				"  		   SELECT tb.fcba, jb.job_forbalancesheetcode, 0 amount, SUM (tb.ending_balance) amount_prev, 0 amount_sbi, 0 amount_prev_sbi" & _
				"              FROM trialbalance tb" & _
				"                   INNER JOIN job jb" & _
				"                      ON (tb.fcba = jb.fcba AND tb.fccode = jb.fccode)" & _
				"             WHERE    period_year = '"& v_txtYear &"' - 1" & _
				"                   AND period_month = '"& v_txtMonth &"'" & _ 
				"          GROUP BY tb.fcba, jb.job_forbalancesheetcode" & _
				"		   UNION ALL" & _
				"			  SELECT tb.fcba, jb.job_forbalancesheetcode, 0 amount, 0 amount_prev, SUM (tb.ending_balance) amount_sbi, 0 amount_prev_sbi" & _
				"              FROM trialbalance tb" & _
				"                   INNER JOIN job jb" & _
				"                      ON (tb.fcba = jb.fcba AND tb.fccode = jb.fccode)" & _
				"             WHERE    period_year = '"& v_txtYear &"'" & _
				"                   AND period_month >= '7'" & _ 
				"          GROUP BY tb.fcba, jb.job_forbalancesheetcode" & _
				" 		   UNION ALL" & _
				"  		   SELECT tb.fcba, jb.job_forbalancesheetcode, 0 amount, 0 amount_prev, 0 amount_sbi, SUM (tb.ending_balance) amount_prev_sbi" & _
				"              FROM trialbalance tb" & _
				"                   INNER JOIN job jb" & _
				"                      ON (tb.fcba = jb.fcba AND tb.fccode = jb.fccode)" & _
				"             WHERE    period_year = '"& v_txtYear &"' - 1" & _
				"                   AND period_month >= '7'" & _ 
				"          GROUP BY tb.fcba, jb.job_forbalancesheetcode) STAGE2" & _
				"            ON (    STAGE1.FCBA = STAGE2.FCBA" & _
				"                AND STAGE1.FCCODE = STAGE2.JOB_FORBALANCESHEETCODE)" & _
				" ORDER BY fccode" 

strPasiva2  = 	"          ,param_1," & _
				"          param_2," & _
				"          fccode," & _
				"          jobbalancesheetcode," & _
				"          NVL (SUM (PASIVA), 0) PASIVA," & _
				"		   NVL (SUM (PASIVA_PREV), 0) PASIVA_PREV," & _
				"          NVL (SUM (PASIVA_SBI), 0) PASIVA_SBI," & _
				"		   NVL (SUM (PASIVA_PREV_SBI), 0) PASIVA_PREV_SBI" 

strPasivaTotal	=	" select GroupPlantation " & selectCons & strPasiva2 &"  From (" & strPasiva & " )" & _ 
                    " " & whereCons & _
                    " group by GroupPlantation" & selectCons & ",param_1, param_2, fccode, jobbalancesheetcode" & _
                    " ORDER BY fccode"

' response.write strAktivaTotal
' response.write strPasivaTotal
Set rsAktiva = Server.CreateObject("ADODB.Recordset")
rsAktiva.Open strAktivaTotal, connstr, adOpenStatic
'

Set rsPasiva = Server.CreateObject("ADODB.Recordset")
rsPasiva.Open strPasivaTotal, connstr, adOpenStatic
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
					   	<div id="header-wrapper">
							<div id="subheader" class="centered">
								<h1>PT. <%=CompanyName%></h1>
								<br>
								<h2>BALANCE SHEET</h2> 
								<br>
								<h3>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h3> 
								<br>
							</div>
						</div>
						<div style="clear:both;"></div>
					</div>
				</th>
			</tr> 
			<!--<div class="printhide">--> c
			<table width="100%" class="xthf-site" style="border-bottom:0px #3A5F0B solid;" >
				<tr>
					<td width="50%">
						<table width="100%" id="product-table" class="xthf-site" style="border-bottom:1px #3A5F0B solid;">
							<tr>
								<th align="center" style="width: 800px;" rowspan="2">AKTIVA</th>
								<th align="center" style="width: 200px;" colspan="2">BULAN INI</th>
								<th align="center" style="width: 200px;" colspan="2">S/D BULAN INI</th>
							</tr> 
							<tr>
								<th align="center" style="width: 200px;"><%=v_txtyear%></th>
								<th align="center" style="width: 200px;"><%=v_txtyear - 1%></th>
								<th align="center" style="width: 200px;"><%=v_txtyear%></th>
								<th align="center" style="width: 200px;"><%=v_txtyear - 1%></th>
							</tr>
							</tr>
								<%
								if rsAktiva.bof and rsAktiva.eof then
								%>
									<tr><td colspan="2">NO DATA AVAILABLE</td></tr>
								<%end if
										v_aktiva = 0
										v_aktiva_prev = 0
										v_aktiva_sbi = 0
										v_aktiva_prev_sbi = 0
										while not rsAktiva.eof
											aktiva = rsAktiva("aktiva")
											param_1= rsAktiva("param_1")
											if NO  mod 2 = 1 then
												cClass = "odd"
											else
												cClass = "default"
											end if
								%>
							<tr> 
								<td style="font-size: 6pt;">
									<strong><i><%=Replace(rsAktiva("jobbalancesheetcode"),"_","&nbsp;")%></i></strong>
								</td>
								<% if cdbl(to_nol(aktiva)) = 0 AND cdbl(to_nol(param_1)) = 0  then %>
									<td align="right"><strong><i></i></strong></td>
									<td align="right"><strong><i></i></strong></td>
									<td align="right"><strong><i></i></strong></td>
									<td align="right"><strong><i></i></strong></td>
								<%else%>
									<td align="right" style="font-size: 6pt;">
										<strong><i><%=FormatNumber(rsAktiva("aktiva"),2)%></i></strong>
									</td>
									<td align="right" style="font-size: 6pt;">
										<strong><i><%=FormatNumber(rsAktiva("aktiva_prev"),2)%></i></strong>
									</td>
									<td align="right" style="font-size: 6pt;">
										<strong><i><%=FormatNumber(rsAktiva("aktiva_sbi"),2)%></i></strong>
									</td>
									<td align="right" style="font-size: 6pt;">
										<strong><i><%=FormatNumber(rsAktiva("aktiva_prev_sbi"),2)%></i></strong>
									</td>
								<%End if%>
							</tr>
								<%	v_aktiva = v_aktiva + cdbl(to_nol(rsAktiva("aktiva")))
									v_aktiva_prev = v_aktiva_prev + cdbl(to_nol(rsAktiva("aktiva_prev")))
									v_aktiva_sbi = v_aktiva_sbi + cdbl(to_nol(rsAktiva("aktiva_sbi")))
									v_aktiva_prev_sbi = v_aktiva_prev_sbi + cdbl(to_nol(rsAktiva("aktiva_prev_sbi")))
									rsAktiva.MoveNext
									Wend
								rsAktiva.Close
								set rsAktiva = Nothing
								%> 
							<tr bgcolor="#F4F4A0">
								<td align="center" style="font-size: 6pt;">
									<strong><i>TOTAL</i></strong>
								</td>
								<td align="right" style="font-size: 6pt;">
									<strong><i><%=formatnumber(cdbl(to_nol(v_aktiva)),2)%></i></strong>
								</td>
								<td align="right" style="font-size: 6pt;">
									<strong><i><%=formatnumber(cdbl(to_nol(v_aktiva_prev)),2)%></i></strong>
								</td>
								<td align="right" style="font-size: 6pt;">
									<strong><i><%=formatnumber(cdbl(to_nol(v_aktiva_sbi)),2)%></i></strong>
								</td>
								<td align="right" style="font-size: 6pt;">
									<strong><i><%=formatnumber(cdbl(to_nol(v_aktiva_prev_sbi)),2)%></i></strong>
								</td>
							</tr> 
						</table>
					</td>
					<td width="50%">
						<table width="100%" id="product-table" class="xthf-site" style="border-bottom:1px #3A5F0B solid;">
							<tr>
								<th align="center" style="width: 800px;" rowspan="2">PASIVA</th>
								<th align="center" style="width: 200px;" colspan="2">BULAN INI</th>
								<th align="center" style="width: 200px;" colspan="2">S/D BULAN INI</th>
							</tr> 
							<tr>
								<th align="center" style="width: 200px;"><%=v_txtyear%></th>
								<th align="center" style="width: 200px;"><%=v_txtyear - 1%></th>
								<th align="center" style="width: 200px;"><%=v_txtyear%></th>
								<th align="center" style="width: 200px;"><%=v_txtyear - 1%></th>
							</tr>
							</tr>
								<%
								if rspasiva.bof and rspasiva.eof then
								%>
									<tr><td colspan="2">NO DATA AVAILABLE</td></tr>
								<%end if
										v_pasiva = 0
										v_pasiva_prev = 0
										v_pasiva_sbi = 0
										v_pasiva_prev_sbi = 0
										while not rspasiva.eof
											pasiva = rspasiva("pasiva")
											param_1= rspasiva("param_1")
											if NO  mod 2 = 1 then
												cClass = "odd"
											else
												cClass = "default"
											end if
								%>
							<tr> 
								<td style="font-size: 6pt;">
									<strong><i><%=Replace(rspasiva("jobbalancesheetcode"),"_","&nbsp;")%></i></strong>
								</td>
								<% if cdbl(to_nol(pasiva)) = 0 AND cdbl(to_nol(param_1)) = 0  then %>
									<td align="right"><strong><i></i></strong></td>
									<td align="right"><strong><i></i></strong></td>
									<td align="right"><strong><i></i></strong></td>
									<td align="right"><strong><i></i></strong></td>
								<%else%>
									<td align="right" style="font-size: 6pt;">
										<strong><i><%=FormatNumber(rspasiva("pasiva"),2)%></i></strong>
									</td>
									<td align="right" style="font-size: 6pt;">
										<strong><i><%=FormatNumber(rspasiva("pasiva_prev"),2)%></i></strong>
									</td>
									<td align="right" style="font-size: 6pt;">
										<strong><i><%=FormatNumber(rspasiva("pasiva_sbi"),2)%></i></strong>
									</td>
									<td align="right" style="font-size: 6pt;">
										<strong><i><%=FormatNumber(rspasiva("pasiva_prev_sbi"),2)%></i></strong>
									</td>
								<%End if%>
							</tr>
								<%	v_pasiva = v_pasiva + cdbl(to_nol(rspasiva("pasiva")))
									v_pasiva_prev = v_pasiva_prev + cdbl(to_nol(rspasiva("pasiva_prev")))
									v_pasiva_sbi = v_pasiva_sbi + cdbl(to_nol(rspasiva("pasiva_sbi")))
									v_pasiva_prev_sbi = v_pasiva_prev_sbi + cdbl(to_nol(rspasiva("pasiva_prev_sbi")))
									rspasiva.MoveNext
									Wend
								rspasiva.Close
								set rspasiva = Nothing
								%> 
							<tr bgcolor="#F4F4A0">
								<td align="center" style="font-size: 6pt;">
									<strong><i>TOTAL</i></strong>
								</td>
								<td align="right" style="font-size: 6pt;">
									<strong><i><%=formatnumber(cdbl(to_nol(v_pasiva)),2)%></i></strong>
								</td>
								<td align="right" style="font-size: 6pt;">
									<strong><i><%=formatnumber(cdbl(to_nol(v_pasiva_prev)),2)%></i></strong>
								</td>
								<td align="right" style="font-size: 6pt;">
									<strong><i><%=formatnumber(cdbl(to_nol(v_pasiva_sbi)),2)%></i></strong>
								</td>
								<td align="right" style="font-size: 6pt;">
									<strong><i><%=formatnumber(cdbl(to_nol(v_pasiva_prev_sbi)),2)%></i></strong>
								</td>
							</tr> 
						</table>
					</td>
				</tr>
			</table>   
	</table>
</body>
</html>






