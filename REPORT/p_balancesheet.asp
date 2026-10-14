<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<% 

v_txtYear		=	Request("txtYear")
v_txtMonth		=	Request("txtMonth")
v_txtdate		=	Request("v_txtdate")
v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")
v_txtOwnership			= Request("txtOwnership")

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

if request("txtYear")="" then
   v_txtYear	=	Year(Session("v_Costbook"))
end if
if request("txtMonth")="" then
   v_txtMonth	=	Month(Session("v_Costbook"))
end if

v_txtdate = v_txtYear &""& v_txtMonth
 

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
				"                   AND period_month >= '7' AND period_month <= '"& v_txtMonth &"'" & _ 
				"          GROUP BY tb.fcba, jb.job_forbalancesheetcode" & _
				" 		   UNION ALL" & _
				"  		   SELECT tb.fcba, jb.job_forbalancesheetcode, 0 amount, 0 amount_prev, 0 amount_sbi, SUM (tb.ending_balance) amount_prev_sbi" & _
				"              FROM trialbalance tb" & _
				"                   INNER JOIN job jb" & _
				"                      ON (tb.fcba = jb.fcba AND tb.fccode = jb.fccode)" & _
				"             WHERE    period_year = '"& v_txtYear &"' - 1" & _
				"                   AND period_month >= '7' AND period_month <= '"& v_txtMonth &"'" & _ 
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
				"                   AND period_month >= '7' AND period_month <= '"& v_txtMonth &"'" & _ 
				"          GROUP BY tb.fcba, jb.job_forbalancesheetcode" & _
				" 		   UNION ALL" & _
				"  		   SELECT tb.fcba, jb.job_forbalancesheetcode, 0 amount, 0 amount_prev, 0 amount_sbi, SUM (tb.ending_balance) amount_prev_sbi" & _
				"              FROM trialbalance tb" & _
				"                   INNER JOIN job jb" & _
				"                      ON (tb.fcba = jb.fcba AND tb.fccode = jb.fccode)" & _
				"             WHERE    period_year = '"& v_txtYear &"' - 1" & _
				"                   AND period_month >= '7' AND period_month <= '"& v_txtMonth &"'" & _ 
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
<html>

<head>
<title>ePlaS Reporting - Balance Sheet</title>
<script language="JavaScript" src="/js/date-picker.js"></script> 
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<link rel="stylesheet" href="/css/jquery-ui.css" />
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<!-- <link rel="stylesheet" href="/css/Neraca.css" type="text/css" /> -->
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
body{font-family:'Arial Narrow';font-size:12px;}
.printhide{display:none;}
}
</style>
</head>

<body>
	<div id="wrapper">
		<!--#include file="../print_excel.asp"-->
		<br><br> 
		<div align="right">
			<a href="javascript:launchwin('print_balancesheet.Asp?period_month=<%=v_txtMonth%>&period_year=<%=v_txtYear%>&Mode=Print','newwindow','width='+(700)+',height='+(505)+',screenX=0,screenY=0,directories=0,fullscreen=0,location=0,menubar=0,scrollbars=1,status=0,toolbar=1')">
				<img src="../IMAGES/Print.gif" width ="25" border="0" title="Print Preview">
				<font face="Verdana" size="1">Print</font>
			</a>&nbsp;&nbsp;
		</div>
		<!--#include file="menutop.asp"-->
		<div class="printhide">
			<table id="form-table" class="centered-margin">
				<form action="<%=ScriptName%>" method="post" name="frmEntry">
					<div class="printhide">
						<!--#include virtual="/REPORT/FormSelectNewSumPNL.html"-->
						<table id="form-table" class="centered-margin">
							<tr>
								<td><font face="verdana" size="1">YEAR</font></td>
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
						<%
							For i = MinYr to MaxYr
						%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
						<%
							Next
						%>
									</select>
								</td>
								<td><font face="verdana" size="1">MONTH</font></td>
								<td width="2">
									<select name="txtMonth">
						<%
								For i = 1 to 15
						%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=BulanAP(i)%></option>
						<%
							Next
						%>
									</select>
								</td>
						<!-- <td>
							<font face="verdana" size="1">OWNERSHIP</font>
						</td>
						<td width="2">
							<select name="txtOwnership" >
								<option value="ALL">--ALL--</option>
								<%
								Dim rsOwnership, strOwnership
								strOwnership	=	" SELECT DISTINCT P.FCCODE,P.FCNAME FROM PARAMETER P WHERE P.PARCODE= 'F0004' AND FCBA = '" & v_txtFcba & "' ORDER BY FCCODE"
								
								Set rsOwnership = Server.CreateObject("ADODB.Recordset")
								rsOwnership.Open strOwnership, connstr, adOpenStatic 
								
								if not rsOwnership.eof then
									Do While Not rsOwnership.eof
									if rsOwnership("FCCODE") = v_txtOwnership then%>
									<option value="<%=rsOwnership("FCCODE")%>" Selected><%=rsOwnership("FCCODE")%></option> 
									<%Else%>
									<option value="<%=rsOwnership("FCCODE")%>"><%=rsOwnership("FCCODE")%></option>  
									<%End if
										rsOwnership.movenext
									loop
								end if
								rsOwnership.Close
								set rsOwnership = Nothing
								%>
							</select>
						</td> -->
								<td><input type="submit" value="View" /></td>
								<td><input type="hidden" name="no-null" value="<%=v_txtnonull%>"></td>
							</tr>
						</table>
					</div>
				</form>
			</table>
		</div>
		
		<div id="header-wrapper">
			<div class="printhide">
			</div>
			<div id="subheader" class="centered">
				<h2>PT. <%=CompanyName%></h2>
				<br>
				<h2>BALANCE SHEET</h2> 
				<br>
				<h3>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h3> 
				<br>
			</div>
		</div> 
	</div>
<table width="100%" class="xthf-site" style="border-bottom:0px #3A5F0B solid;" >
	<tr>
		<td width="50%" valign="top">
			<table width="100%" id="product-table" class="xthf-site" style="border-bottom:1px #3A5F0B solid;">
				<tr>
					<th align="center" style="width: 800px;" rowspan="2">AKTIVA</th>
					<th align="center" style="width: 200px;" colspan="2">BULAN INI</th>
				</tr> 
				<tr>
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
					<%else%>
						<td align="right" style="font-size: 6pt;">
							<strong><i><%=FormatNumber(rsAktiva("aktiva"),2)%></i></strong>
						</td>
						<td align="right" style="font-size: 6pt;">
							<strong><i><%=FormatNumber(rsAktiva("aktiva_prev"),2)%></i></strong>
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
				</tr> 
			</table>
		</td>
		<td width="50%" valign="top">
			<table width="100%" id="product-table" class="xthf-site" style="border-bottom:1px #3A5F0B solid;">
				<tr>
					<th align="center" style="width: 800px;" rowspan="2">PASIVA</th>
					<th align="center" style="width: 200px;" colspan="2">BULAN INI</th>
				</tr> 
				<tr>
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
					<%else%>
						<td align="right" style="font-size: 6pt;">
							<strong><i><%=FormatNumber(rspasiva("pasiva"),2)%></i></strong>
						</td>
						<td align="right" style="font-size: 6pt;">
							<strong><i><%=FormatNumber(rspasiva("pasiva_prev"),2)%></i></strong>
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
				<tr><td style="font-size: 6pt;">&nbsp;</td><td style="font-size: 6pt;">&nbsp;</td><td style="font-size: 6pt;">&nbsp;</td></tr>
				<tr><td style="font-size: 6pt;">&nbsp;</td><td style="font-size: 6pt;">&nbsp;</td><td style="font-size: 6pt;">&nbsp;</td></tr>
				<tr><td style="font-size: 6pt;">&nbsp;</td><td style="font-size: 6pt;">&nbsp;</td><td style="font-size: 6pt;">&nbsp;</td></tr>
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
				</tr> 
			</table>
		</td>
	</tr>
</table>  
</body>
</html>
