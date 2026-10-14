<%@Language=VBScript%>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtYear, v_txtMonth
Dim ScriptName, NO, cClass
Dim nSALDOAWAL, nDEBET, nKREDIT, nSALDOAKHIR
Dim v_txtJobCode,v_txtJobCodeto 

v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")
v_Isactive              = Request("txtIsactive")

v_txtYear		=	Request("txtYear")
v_txtMonth		=	Request("txtMonth")

v_txtjobcode        = Request("txtJobCode")
v_txtjobcodeto      = Request("txtJobCodeto")

If Len(Trim(v_txtcompany)) = 0 Then
    v_txtcompany = Session("Companycode")
End If

Dim selectFields, strFieldCons, selectCons, whereCons, GroupbyCons, orderCons

strFieldCons	=	" USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation" & _
					" , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation" & _
					" , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion" & _
					" , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region" & _
					" , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company" & _
					" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba"
					
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
		var_colspan	=	var_colspan	+ 1
	Else
		selectCons	=	selectCons	& " , company"
		whereCons	=	whereCons	& " And company = '" & v_txtCompany & "'"
		orderCons	=	orderCons	& " , COMPANY_API.get_OrderData(company)"
		var_colspan	=	var_colspan	+ 1
		
		q	=	" SELECT FCCODE" & _
				" , FCCOMPANYCODE" & _
				" , FCNAME" & _
				" , FCCOMPANYNAME" & _
				" , FCMANAGER" & _
				" , FCKTU" & _
				" , FCCROP" & _
				" , FCTYPE" & _
				" FROM BUSINESSUNIT WHERE Isactive='TRUE' AND " & _
				" FCCOMPANYCODE = '" & v_txtCompany & "'"	
		set rsCat = connStr.Execute(q)		
		if not rsCat.eof then
			CompanyName = rsCat("FCCOMPANYNAME")	
		end if
		q	=	""
		rsCat.Close
		set rsCat = Nothing
	End If
End If

If Len(Trim(v_txtfcba)) = 0 Then v_txtfcba = "--SUMMARY--"

If Len(Trim(v_txtfcba)) > 0 Then
	if v_txtfcba = "--SUMMARY--" then
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , nourut"
		nSpanTotal	=	nSpanTotal	- 1
	Elseif v_txtfcba = "--ALL--" then
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	wherecons	& " "
		orderCons	=	orderCons	& " , nourut"
	Else
		selectCons	=	selectCons	& " , nourut, fcba"
		whereCons	=	whereCons	& " And fcba = '" & v_txtfcba & "'"
		orderCons	=	orderCons	& " , nourut"
		
		q	=	" SELECT FCCODE" & _
				" , FCCOMPANYCODE" & _
				" , FCNAME" & _
				" , FCCOMPANYNAME" & _
				" , FCMANAGER" & _
				" , FCKTU" & _
				" , FCCROP" & _
				" , FCTYPE" & _
				" FROM BUSINESSUNIT WHERE Isactive='TRUE' AND " & _
				" FCCODE = '" & v_txtfcba & "'"	
		set rsCat = connStr.Execute(q)		
		if not rsCat.eof then
			fcbaname = rsCat("FCNAME")
			CompanyName=rsCat("FCCOMPANYNAME")
			MgrName = rsCat("FCMANAGER")
			KTUName = rsCat("FCKTU")
			CropName=rsCat("FCCROP")
			FcbaType=rsCat("FCTYPE")
		end if
		q	=	""
		rsCat.Close
		set rsCat = Nothing
	End If
End If

If v_Isactive <> "ALL" then
   whereCons	=	whereCons	& " and businessunit_api.Isactive(fcba) = '" & v_Isactive & "'"	
end if

v_txtYear = Request("txtYear")
v_txtMonth = Request("txtMonth")

If "" & v_txtYear = "" Then v_txtYear = Year(Session("v_Costbook"))
If "" & v_txtMonth = "" Then v_txtMonth = Month(Session("v_Costbook"))


Dim rsNeracaTemp, strNeracaTemp



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
	varFcbaName = "ALL UNIT"
	varCompanyName = Session("COMPANY")
	varCompanyCode = Session("COMPANYCODE")
End If
rsBaInfo.Close
set rsBaInfo = Nothing


'--New Query--
IF v_txtCompany <> "--ALL--" Then
		IF v_txtFcba<> "--ALL--" Then 
		strNeracaTemp	=	" SELECT" & _
							"     TEMPDATA.FCBA" & _
							"     , NYEAR" & _
							"     , NMONTH" & _
							"     , JOBCODE" & _
							"     , FCNAME JOBNAME" & _
							"     , SUM(SALDOAWAL)SALDOAWAL" & _
							"     , SUM(DEBET) DEBET" & _
							"     , SUM(KREDIT) KREDIT" & _
							"     , SUM(SALDOAWAL)+SUM(DEBET)+SUM(KREDIT) SALDOAKHIR " & _
							" FROM (" & _
							"     SELECT" & _
							"         FCBA" & _
							"         , " & v_txtYear & " NYEAR" & _
							"         , " & v_txtmonth & " NMONTH" & _
							"         , FCCODE JOBCODE" & _
							"         , 0 SALDOAWAL" & _
							"         , 0 DEBET" & _
							"         , 0 KREDIT" & _
							"     FROM JOB" & _
							"     WHERE ACTIVATION = 'Y' AND FCBA = '" & v_txtFcba & "'" & _
							"     UNION ALL" & _
							"     SELECT" & _
							"         FCBA" & _
							"         , PERIOD_YEAR NYEAR" & _
							"         , PERIOD_MONTH NMONTH" & _
							"         , FCCODE JOBCODE" & _
							"         , OPENING_BALANCE SALDOAWAL" & _
							"         , 0 DEBET" & _
							"         , 0 KREDIT" & _
							"     FROM TRIALBALANCE" & _
							"     WHERE FCBA = '" & v_txtFcba & "' AND PERIOD_YEAR = " & v_txtYear & " AND PERIOD_MONTH = " & v_txtmonth & _
							"     UNION ALL" & _
							"     SELECT" & _
							"         FCBA" & _
							"         , accountingyear NYEAR" & _
							"         , accountingperiod NMONTH" & _
							"         , JOBCODE" & _
							"         , 0 SALDOAWAL" & _
							"         , AMOUNT DEBET" & _
							"         , 0 KREDIT" & _
							"     FROM COST_ACTIVITY" & _
							"     WHERE FCBA = '" & v_txtFcba & "' AND accountingyear= " & v_txtYear & " AND accountingperiod = " & v_txtmonth & " AND AMOUNT > 0" & _
							"     UNION ALL" & _
							"     SELECT" & _
							"         FCBA" & _
							"         , accountingyear NYEAR" & _
							"         , accountingperiod NMONTH" & _
							"         , JOBCODE" & _
							"         , 0 SALDOAWAL" & _
							"         , 0 DEBET" & _
							"         , AMOUNT KREDIT" & _
							"     FROM COST_ACTIVITY" & _
							"     WHERE FCBA = '" & v_txtFcba & "' AND accountingyear = " & v_txtYear & " AND accountingperiod  = " & v_txtmonth & "  AND AMOUNT < 0" & _
							" ) TEMPDATA, JOB" & _
							" WHERE (TEMPDATA.FCBA = JOB.FCBA AND TEMPDATA.JOBCODE = JOB.FCCODE AND JOB.ACTIVATION='Y')"
			if LEN(TRIM(v_txtjobcode)) > 0 AND LEN(TRIM(v_txtjobcodeto)) > 0 Then
				strNeracaTemp	=	strNeracaTemp & " and (jobcode between '" & trim(v_txtjobCode) & "' and '" & trim(v_txtjobCodeto) & "') " 
				else
				strNeracaTemp	=	strNeracaTemp & " and jobcode like '" & trim(v_txtjobCode) & "%'"
			end if

			strNeracaTemp	=	strNeracaTemp & " GROUP BY TEMPDATA.FCBA, NYEAR, NMONTH, JOBCODE, FCNAME"
			strNeracaTemp	=	strNeracaTemp & " ORDER BY FCBA, NYEAR, NMONTH, JOBCODE"					
		Else
			strNeracaTemp	=	" SELECT" & _
							"      NYEAR" & _
							"     , NMONTH" & _
							"     , JOBCODE" & _
							"     , FCNAME JOBNAME" & _
							"     , SUM(SALDOAWAL)SALDOAWAL" & _
							"     , SUM(DEBET) DEBET" & _
							"     , SUM(KREDIT) KREDIT" & _
							"     , SUM(SALDOAWAL)+SUM(DEBET)+SUM(KREDIT) SALDOAKHIR " & _
							" FROM (" & _
							"     SELECT" & _
							"         FCBA" & _
							"         , " & v_txtYear & " NYEAR" & _
							"         , " & v_txtmonth & " NMONTH" & _
							"         , FCCODE JOBCODE" & _
							"         , 0 SALDOAWAL" & _
							"         , 0 DEBET" & _
							"         , 0 KREDIT" & _
							"     FROM JOB" & _
							"     WHERE ACTIVATION = 'Y' AND  FCBA in (SELECT FCCODE FROM  BUSINESSUNIT WHERE Isactive='TRUE' AND FCCOMPANYCODE='" & v_txtCompany & "')" & _
							"     UNION ALL" & _
							"     SELECT" & _
							"         FCBA" & _
							"         , PERIOD_YEAR NYEAR" & _
							"         , PERIOD_MONTH NMONTH" & _
							"         , FCCODE JOBCODE" & _
							"         , OPENING_BALANCE SALDOAWAL" & _
							"         , 0 DEBET" & _
							"         , 0 KREDIT" & _
							"     FROM TRIALBALANCE" & _
							"     WHERE FCBA in (SELECT FCCODE FROM  BUSINESSUNIT WHERE Isactive='TRUE' AND FCCOMPANYCODE='" & v_txtCompany & "') AND PERIOD_YEAR = " & v_txtYear & " AND PERIOD_MONTH = " & v_txtmonth & _
							"     UNION ALL" & _
							"     SELECT" & _
							"         FCBA" & _
							"         , accountingyear NYEAR" & _
							"         , accountingperiod NMONTH" & _
							"         , JOBCODE" & _
							"         , 0 SALDOAWAL" & _
							"         , AMOUNT DEBET" & _
							"         , 0 KREDIT" & _
							"     FROM COST_ACTIVITY" & _
							"     WHERE FCBA in (SELECT FCCODE FROM  BUSINESSUNIT WHERE Isactive='TRUE' AND FCCOMPANYCODE='" & v_txtCompany & "') AND accountingyear= " & v_txtYear & " AND accountingperiod = " & v_txtmonth & " AND AMOUNT > 0" & _
							"     UNION ALL" & _
							"     SELECT" & _
							"         FCBA" & _
							"         , accountingyear NYEAR" & _
							"         , accountingperiod NMONTH" & _
							"         , JOBCODE" & _
							"         , 0 SALDOAWAL" & _
							"         , 0 DEBET" & _
							"         , AMOUNT KREDIT" & _
							"     FROM COST_ACTIVITY" & _
							"     WHERE FCBA in (SELECT FCCODE FROM BUSINESSUNIT WHERE Isactive='TRUE' AND FCCOMPANYCODE='" & v_txtCompany & "') AND accountingyear = " & v_txtYear & " AND accountingperiod  = " & v_txtmonth & "  AND AMOUNT < 0" & _
							" ) TEMPDATA, JOB" & _
							" WHERE (TEMPDATA.FCBA = JOB.FCBA AND TEMPDATA.JOBCODE = JOB.FCCODE)" 

			if LEN(TRIM(v_txtjobcode)) > 0 AND LEN(TRIM(v_txtjobcodeto)) > 0 Then
			strNeracaTemp	=	strNeracaTemp & " and (jobcode between '" & trim(v_txtjobCode) & "' and '" & trim(v_txtjobCodeto) & "') " 
			else
			strNeracaTemp	=	strNeracaTemp & " and jobcode like '" & trim(v_txtjobCode) & "%'"
			end if

			strNeracaTemp	=	strNeracaTemp & " GROUP BY NYEAR, NMONTH, JOBCODE, FCNAME"
			strNeracaTemp	=	strNeracaTemp & " ORDER BY NYEAR, NMONTH, JOBCODE"	
		End IF
Else
		strNeracaTemp	=	" SELECT" & _
							"      NYEAR" & _
							"     , NMONTH" & _
							"     , JOBCODE" & _
							"     , FCNAME JOBNAME" & _
							"     , SUM(SALDOAWAL)SALDOAWAL" & _
							"     , SUM(DEBET) DEBET" & _
							"     , SUM(KREDIT) KREDIT" & _
							"     , SUM(SALDOAWAL)+SUM(DEBET)+SUM(KREDIT) SALDOAKHIR " & _
							" FROM (" & _
							"     SELECT" & _
							"         FCBA" & _
							"         , " & v_txtYear & " NYEAR" & _
							"         , " & v_txtmonth & " NMONTH" & _
							"         , FCCODE JOBCODE" & _
							"         , 0 SALDOAWAL" & _
							"         , 0 DEBET" & _
							"         , 0 KREDIT" & _
							"     FROM JOB" & _
							"     WHERE ACTIVATION = 'Y' AND  FCBA in (SELECT FCCODE FROM  BUSINESSUNIT WHERE Isactive='TRUE' AND FCCOMPANYCODE='" & v_txtCompany & "')" & _
							"     UNION ALL" & _
							"     SELECT" & _
							"         FCBA" & _
							"         , PERIOD_YEAR NYEAR" & _
							"         , PERIOD_MONTH NMONTH" & _
							"         , FCCODE JOBCODE" & _
							"         , OPENING_BALANCE SALDOAWAL" & _
							"         , 0 DEBET" & _
							"         , 0 KREDIT" & _
							"     FROM TRIALBALANCE" & _
							"     WHERE FCBA in (SELECT FCCODE FROM  BUSINESSUNIT WHERE Isactive='TRUE' AND FCCOMPANYCODE='" & v_txtCompany & "') AND PERIOD_YEAR = " & v_txtYear & " AND PERIOD_MONTH = " & v_txtmonth & _
							"     UNION ALL" & _
							"     SELECT" & _
							"         FCBA" & _
							"         , accountingyear NYEAR" & _
							"         , accountingperiod NMONTH" & _
							"         , JOBCODE" & _
							"         , 0 SALDOAWAL" & _
							"         , AMOUNT DEBET" & _
							"         , 0 KREDIT" & _
							"     FROM COST_ACTIVITY" & _
							"     WHERE FCBA in (SELECT FCCODE FROM  BUSINESSUNIT WHERE Isactive='TRUE' AND FCCOMPANYCODE='" & v_txtCompany & "') AND accountingyear= " & v_txtYear & " AND accountingperiod = " & v_txtmonth & " AND AMOUNT > 0" & _
							"     UNION ALL" & _
							"     SELECT" & _
							"         FCBA" & _
							"         , accountingyear NYEAR" & _
							"         , accountingperiod NMONTH" & _
							"         , JOBCODE" & _
							"         , 0 SALDOAWAL" & _
							"         , 0 DEBET" & _
							"         , AMOUNT KREDIT" & _
							"     FROM COST_ACTIVITY" & _
							"     WHERE FCBA in (SELECT FCCODE FROM BUSINESSUNIT WHERE Isactive='TRUE' AND FCCOMPANYCODE='" & v_txtCompany & "') AND accountingyear = " & v_txtYear & " AND accountingperiod  = " & v_txtmonth & "  AND AMOUNT < 0" & _
							" ) TEMPDATA, JOB" & _
							" WHERE (TEMPDATA.FCBA = JOB.FCBA AND TEMPDATA.JOBCODE = JOB.FCCODE)" 

			if LEN(TRIM(v_txtjobcode)) > 0 AND LEN(TRIM(v_txtjobcodeto)) > 0 Then
			strNeracaTemp	=	strNeracaTemp & " and (jobcode between '" & trim(v_txtjobCode) & "' and '" & trim(v_txtjobCodeto) & "') " 
			else
			strNeracaTemp	=	strNeracaTemp & " and jobcode like '" & trim(v_txtjobCode) & "%'"
			end if

			strNeracaTemp	=	strNeracaTemp & " GROUP BY NYEAR, NMONTH, JOBCODE, FCNAME"
			strNeracaTemp	=	strNeracaTemp & " ORDER BY NYEAR, NMONTH, JOBCODE"	
END If

'Response.Write "|" & strNeracaTemp & "|" 

Set rsNeracaTemp = Server.CreateObject("ADODB.Recordset")
rsNeracaTemp.Open strNeracaTemp, connstr, adOpenStatic

%>
<html>

<head>
<title>ePlaS Reporting - Neraca Sementara</title>
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
				<!--#include virtual="/REPORT/FormSelectNewSUMAll.html"-->
				<table id="form-table" class="centered-margin" >	
						<td>
							<font face="verdana" size="1">TAHUN</font>
						</td>
						<% 
							Dim mon, Yr, MinYr, MaxYr, i, MM
							mon = Month(Session("v_Costbook"))
							If len(v_txtMonth)=0 then v_txtMonth=mon
						
							Yr = Year(Session("v_Costbook"))
							MinYr = Yr - 2
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
						<td>
							<font face="verdana" size="1">BULAN</font>
						</td>
						<td width="2">
							<select name="txtMonth">
							<%
							Dim rsBln, strBln
							strBln	=	"SELECT DISTINCT accountingperiod FCCODE, DESCRIPTION FCNAME FROM accountingperiod WHERE FCBA IN (SELECT FCCODE FROM BUSINESSUNIT WHERE FCCOMPANYCODE='" & v_txtCompany & "') ORDER BY FCCODE ASC"
							
							Set rsBln = Server.CreateObject("ADODB.Recordset")
							rsBln.Open strBln, connstr, adOpenStatic 
							
							if not rsBln.eof then
							Do While Not rsBln.eof
							if cint(rsBln("FCCODE")) = cint(v_txtMonth) then%>
									<option value="<%=rsBln("FCCODE")%>" Selected><%=rsBln("FCCODE")%> - <%=rsBln("FCNAME")%></option> 
								 <%Else%>
									<option value="<%=rsBln("FCCODE")%>"><%=rsBln("FCCODE")%> - <%=rsBln("FCNAME")%></option>  
							<%End if
								rsBln.movenext
							loop
							end if
							rsBln.Close
							set rsBln = Nothing
							%>
							</select>
						</td>
				</table>
				<table class="centered-margin" >	
						<tr>
							<td><font face="verdana" size="1">JOBCODE</font></td>
							<td>
							<select name="txtJobCode" class="chosen-select">
							<%
							Dim rsJobCode, strJobCode
							strJobCode	=	"SELECT ' ' FCCODE,'--ALL--' FCNAME FROM DUAL UNION SELECT DISTINCT FCCODE, FCNAME FROM JOB WHERE FCBA IN (SELECT FCCODE FROM BUSINESSUNIT WHERE FCCOMPANYCODE='" & v_txtCompany & "') ORDER BY FCCODE ASC"
							
							Set rsJobCode = Server.CreateObject("ADODB.Recordset")
							rsJobCode.Open strJobCode, connstr, adOpenStatic 
							
							if not rsJobCode.eof then
							Do While Not rsJobCode.eof
							if trim(rsJobCode("FCCODE")) = v_txtJobCode then%>
									<option value="<%=rsJobCode("FCCODE")%>" Selected><%=rsJobCode("FCCODE")%> - <%=rsJobCode("FCNAME")%></option> 
								 <%Else%>
									<option value="<%=rsJobCode("FCCODE")%>"><%=rsJobCode("FCCODE")%> - <%=rsJobCode("FCNAME")%></option>  
							<%End if
								rsJobCode.movenext
							loop
							end if
							rsJobCode.Close
							set rsJobCode = Nothing
							%>
							</select>
							</td>
							<td><font face="verdana" size="1">TO</font></td>
							<td>
							<select name="txtJobCodeto" class="chosen-select">
							<%
							
							strJobCode	=	"SELECT ' ' FCCODE,'--ALL--' FCNAME FROM DUAL UNION SELECT DISTINCT FCCODE, FCNAME FROM JOB WHERE FCBA IN (SELECT FCCODE FROM BUSINESSUNIT WHERE FCCOMPANYCODE='" & v_txtCompany & "') ORDER BY FCCODE ASC"
							
							Set rsJobCode = Server.CreateObject("ADODB.Recordset")
							rsJobCode.Open strJobCode, connstr, adOpenStatic 
							
							if not rsJobCode.eof then
							Do While Not rsJobCode.eof
							if trim(rsJobCode("FCCODE")) = v_txtJobCodeto then%>
									<option value="<%=rsJobCode("FCCODE")%>" Selected><%=rsJobCode("FCCODE")%> - <%=rsJobCode("FCNAME")%></option> 
								 <%Else%>
									<option value="<%=rsJobCode("FCCODE")%>"><%=rsJobCode("FCCODE")%> - <%=rsJobCode("FCNAME")%></option>  
							<%End if
							rsJobCode.movenext
							loop
							end if
							rsJobCode.Close
							set rsJobCode = Nothing
							%>
							</select>
							</td>
						</tr>
						</table>
					</tr>
					<tr>
						<td><input type="submit" value="View" /></td>
					</tr>	
				</table>
			</div>
			</form>
</div>
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="50" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
			<div id="head_judul">
				<% if  CompanyName <>  "" then %>
					<h4>PT.<%=CompanyName%></h4> 
				<% End if%>
				<% If  FcbaName <> ""  then %>
					   <h5><%=FcbaName%></h5> 
				<% Else%>
				       <h5>All UNIT <%=v_txtCompany%></h5> 
				<% End if%>
			</div>
			<div id="subheader" class="centered" style="width:50%; margin:auto;">
				<h2>NERACA SEMENTARA</h3>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=BulanAP(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
			</div>
			</div>
</th>
  </tr>
  
	<tr>
			<th style="background-color:#00A300;color:#fff;">NO</th>
			<th style="background-color:#00A300;color:#fff;">KODE</th>
			<th style="background-color:#00A300;color:#fff;">PERKIRAAN</th>
			<th style="background-color:#00A300;color:#fff;">SALDO AWAL</th>
			<th style="background-color:#00A300;color:#fff;">DEBET</th>
			<th style="background-color:#00A300;color:#fff;">KREDIT</th>
			<th style="background-color:#00A300;color:#fff;">SALDO AKHIR</th>
		</tr>
</thead>
<tbody>
	<%
	if rsNeracaTemp.bof and rsNeracaTemp.eof then
	%>
		<tr><td colspan="7">NO DATA AVAILABLE</td></tr>
	<%end if
			NO = 1
			while not rsNeracaTemp.eof
				if NO  mod 2 = 1 then
					cClass = "odd"
				else
					cClass = "default"
				end if
	%>
		<tr class="<%=cClass%>">
			<td align="center"><%=NO%></td>
			<td align="center"><%=rsNeracaTemp("JOBCODE")%></td>
			<td align="left"><%=rsNeracaTemp("JOBNAME")%></td>
			<td align="right"><%=FormatNumber(rsNeracaTemp("SALDOAWAL"),2)%></td>
			<td align="right"><%=FormatNumber(rsNeracaTemp("DEBET"),2)%></td>
			<td align="right"><%=FormatNumber(rsNeracaTemp("KREDIT"),2)%></td>
			<td align="right"><%=FormatNumber(rsNeracaTemp("SALDOAKHIR"),2)%></td>
		</tr>
	<%
				NO = NO + 1
				nSALDOAWAL =  nSALDOAWAL + Cdbl(rsNeracaTemp("SALDOAWAL"))
				nDEBET =  nDEBET + Cdbl(rsNeracaTemp("DEBET"))
				nKREDIT =  nKREDIT + Cdbl(rsNeracaTemp("KREDIT"))
				nSALDOAKHIR =  nSALDOAKHIR+ Cdbl(rsNeracaTemp("SALDOAKHIR"))
				rsNeracaTemp.MoveNext
			Wend
	rsNeracaTemp.Close
	set rsNeracaTemp = Nothing
	%>
		<tr bgcolor="#F4F4A0">
			<td align="center" colspan="3"><strong>TOTAL</strong></td>
			<td align="right"><strong><%=FormatNumber(nSALDOAWAL,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(nDEBET,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(nKREDIT,2)%></strong></td>
			<td align="right"><strong><%=FormatNumber(nSALDOAKHIR,2)%></strong></td>
		</tr>
        </tbody>
	</table>

	<table width="90%" class="data">
		<tr> 
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td width="50%" align="center"><font face="Verdana" size="1">Approved by,</font></td>
			<td width="50%" align="center"><font face="Verdana" size="1">Prepared by,</font></td>
		</tr>
		<tr> 
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr> 
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr> 
			<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("FCMANAGER")%></u></strong></td>
			<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("KTU")%></u></strong></td>
		</tr>
		<tr> 
			<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
			<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
		</tr>
	</table>
</div>
</body>
</html>