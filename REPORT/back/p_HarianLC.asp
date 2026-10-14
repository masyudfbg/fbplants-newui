<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->

<%
Dim v_txtYear, v_txtMonth, v_txtJobcode, v_txtAfdeling, v_txtTahunTanam, v_txtBlok
Dim ScriptName
Dim rsHarianLC, strHarianLC, strMain, strHarianLC2
Dim NO, cClass, nday, xNum, nSpanTotal

nSpanTotal = 9

v_txtGroupPlantation 	= Request("txtGroupPlantation")
v_txtPlantation  		= Request("txtPlantation")
v_txtGroupRegion  		= Request("txtGroupRegion")
v_txtRegion  			= Request("txtRegion")
v_txtcompany  			= Request("txtCompany")
v_txtfcba 				= Request("txtfcba")

Dim selectFields, strFieldCons, selectCons, whereCons, GroupbyCons, OrderbyCons

strFieldCons	=	" USER_COMPANY_API.GetGroupPlantationCode('" & Session("CodeName") & "',fcba) GroupPlantation" & _
					" , USER_COMPANY_API.GetPlantationCode('" & Session("CodeName") & "',fcba) Plantation" & _
					" , USER_COMPANY_API.GetGroupRegionCode('" & Session("CodeName") & "',fcba) groupregion" & _
					" , USER_COMPANY_API.GetRegionCode('" & Session("CodeName") & "',fcba) region" & _
					" , USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fcba) company" & _
					" , BUSINESSUNIT_API.get_OrderData(fcba) NoUrut, fcba"
					
wherecons	=	" where company is not null "


If Len(Trim(v_txtGroupPlantation)) > 0 Then
	if v_txtGroupPlantation = "--ALL--" then
		whereCons = wherecons & " "
	Elseif v_txtGroupPlantation = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		whereCons = wherecons & " And GroupPlantation = '" & v_txtGroupPlantation & "' "
	end if
End If

If Len(Trim(v_txtPlantation)) > 0 Then
	if v_txtPlantation = "--ALL--" then
		selectCons = selectCons & " , Plantation"
		whereCons = wherecons & " "
	Elseif v_txtPlantation = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , Plantation"
		whereCons = whereCons & " And  Plantation = '" & v_txtPlantation & "' "
	end if
End If

If Len(Trim(v_txtGroupRegion)) > 0 Then
	if v_txtGroupRegion = "--ALL--" then
		selectCons = selectCons & " , groupregion"
		whereCons = wherecons & " "
	Elseif v_txtGroupRegion = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , groupregion"
		whereCons = whereCons & " And  GroupRegion = '" & v_txtGroupRegion & "' "
	end if
End If

If Len(Trim(v_txtRegion)) > 0 Then
	if v_txtRegion = "--ALL--" then
		selectCons = selectCons & " , region"
		whereCons = wherecons & " "
	Elseif v_txtRegion = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , region"
		whereCons = whereCons & " And  Region = '" & v_txtRegion & "' "
	end if
End If

If Len(Trim(v_txtCompany)) > 0 Then
	if v_txtCompany = "--ALL--" then
		selectCons = selectCons & " , company"
		whereCons = wherecons & " "
	Elseif v_txtCompany = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , company"
		whereCons = whereCons & " And  company = '" & v_txtCompany & "' "
		q = " Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit" & _
			" where FCCOMPANYCODE = '" & v_txtCompany & "'"	
		set rsCat = connStr.Execute(q)
		if not rsCat.eof then
		   CompanyName=rsCat("FCCOMPANYNAME")			   
		end if
	end if
End If

If Len(Trim(v_txtfcba)) > 0 Then
	if v_txtfcba = "--ALL--" then
		selectCons = selectCons & " , NoUrut, FCBA"
		whereCons = wherecons & " "
	Elseif v_txtfcba = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , NoUrut, FCBA"
		whereCons = whereCons & " And  fcba = '" & v_txtfcba & "' "
		dim q, FcbaName, CompanyName, MgrName, KTUName, CropName, FcbaType
		q = "Select fccode, fccompanycode, FCNAME, FCCOMPANYNAME, FCMANAGER, FCKTU, FCCROP, FCTYPE from BusinessUnit where fcCode = '" & v_txtfcba & "'"
		set rsCat = connStr.Execute(q)
		if not rsCat.eof then	
		   fcbaname		= rsCat("fcName")
		   CompanyName	= rsCat("FCCOMPANYNAME")
		   MgrName		= rsCat("FCMANAGER")
		   KTUName		= rsCat("FCKTU")
		   CropName		= rsCat("FCCROP")
		   FcbaType		= rsCat("fcType")
		end if
	end if
End If

v_txtYear	=	Request("txtYear")
v_txtMonth	=	Request("txtMonth")

If v_txtYear = "" Then
	v_txtYear	=	Year(Session("v_Costbook"))
End If

If v_txtMonth = "" Then
	v_txtMonth	=	Month(Session("v_Costbook"))
End If

Dim varLastDayofMonth : varLastDayofMonth = v_txtMonth & "/" & GetLastDay(v_txtMonth,v_txtYear) & "/" & v_txtYear

v_txtJobcode	=	Request("txtJobcode")
v_txtAfdeling	=	Request("txtAfdeling")
v_txtTahunTanam	=	Request("txtTahunTanam")
v_txtBlok		=	Request("txtBlok")

If v_txtJobcode = "" Then
	v_txtJobcode	=	"--ALL--"
End If

If Len(Trim(v_txtJobcode)) > 0 Then
	if v_txtJobcode = "--ALL--" then
		whereCons = wherecons & " "
	Elseif v_txtJobcode = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		whereCons = whereCons & " And  JOBCODE = '" & v_txtJobcode & "' "
	end if
End If

If Len(Trim(v_txtAfdeling)) > 0 Then
	if v_txtAfdeling = "--ALL--" then
		selectCons = selectCons & " , AFDELING"
		whereCons = wherecons & " "
	Elseif v_txtAfdeling = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , AFDELING"
		whereCons = whereCons & " And  AFDELING = '" & v_txtAfdeling & "' "
	end if
End If

If Len(Trim(v_txtTahunTanam)) > 0 Then
	if v_txtTahunTanam = "--ALL--" then
		selectCons = selectCons & " , TAHUN_TANAM"
		whereCons = wherecons & " "
	Elseif v_txtTahunTanam = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , TAHUN_TANAM"
		whereCons = whereCons & " And  TAHUN_TANAM = '" & v_txtTahunTanam & "' "
	end if
End If

If Len(Trim(v_txtBlok)) > 0 Then
	if v_txtBlok = "--ALL--" then
		selectCons = selectCons & " , Blok"
		whereCons = wherecons & " "
	Elseif v_txtBlok = "--SUMMARY--" then
		whereCons = wherecons & " "
	Else
		selectCons = selectCons & " , Blok"
		whereCons = whereCons & " And  Blok = '" & v_txtBlok & "' "
	end if
End If

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strMain		=	" SELECT" & strFieldCons & _
				"          , JOBCODE," & _
				"          JOB_API.GET_JOBCODE_NAME (JOBCODE, FCBA) JOBNAME," & _
				"          AFDELING," & _
				"          TAHUN_TANAM," & _
				"          BLOK," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 1, QTY, 0)) DAY_1," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 2, QTY, 0)) DAY_2," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 3, QTY, 0)) DAY_3," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 4, QTY, 0)) DAY_4," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 5, QTY, 0)) DAY_5," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 6, QTY, 0)) DAY_6," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 7, QTY, 0)) DAY_7," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 8, QTY, 0)) DAY_8," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 9, QTY, 0)) DAY_9," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 10, QTY, 0)) DAY_10," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 11, QTY, 0)) DAY_11," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 12, QTY, 0)) DAY_12," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 13, QTY, 0)) DAY_13," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 14, QTY, 0)) DAY_14," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 15, QTY, 0)) DAY_15," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 16, QTY, 0)) DAY_16," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 17, QTY, 0)) DAY_17," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 18, QTY, 0)) DAY_18," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 19, QTY, 0)) DAY_19," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 20, QTY, 0)) DAY_20," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 21, QTY, 0)) DAY_21," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 22, QTY, 0)) DAY_22," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 23, QTY, 0)) DAY_23," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 24, QTY, 0)) DAY_24," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 25, QTY, 0)) DAY_25," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 26, QTY, 0)) DAY_26," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 27, QTY, 0)) DAY_27," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 28, QTY, 0)) DAY_28," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 29, QTY, 0)) DAY_29," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 30, QTY, 0)) DAY_30," & _
				"          SUM (DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'DD')), 31, QTY, 0)) DAY_31," & _
				"          SUM (" & _
				"             DECODE (TO_NUMBER (TO_CHAR (TANGGAL, 'YYYY'))," & _
				"                     " & v_txtYear & ", QTY," & _
				"                     TO_NUMBER (TO_CHAR (TANGGAL, 'MM')), " & v_txtMonth & "," & _
				"                     QTY, 0))" & _
				"             REAL," & _
				"          NVL (BUDGET_TRANSACTION_ESTATE_API.GETOUTPUT_SHI (" & _
				"                  FCBA," & _
				"                  " & v_txtYear & "," & _
				"                  1," & _
				"                  1," & _
				"                  ''," & _
				"                  JOBCODE," & _
				"                  ''," & _
				"                  'FISIK'," & _
				"                  ''," & _
				"                  BLOK," & _
				"                  AFDELING," & _
				"                  TAHUN_TANAM," & _
				"                  ''," & _
				"                  TO_DATE ('" & varLastDayofMonth & "', 'MM/DD/YYYY'))," & _
				"               0)" & _
				"             BUDGET" & _
				"     FROM (  SELECT FCBA," & _
				"                    TANGGAL," & _
				"                    JOBCODE," & _
				"                    AFDELING," & _
				"                    TAHUN_TANAM," & _
				"                    BLOK," & _
				"                    SUM (QTY) QTY" & _
				"               FROM (SELECT JJ.FCBA," & _
				"                            TO_DATE ('" & varLastDayofMonth & "', 'MM/DD/YYYY') TANGGAL," & _
				"                            JJ.FCCODE JOBCODE," & _
				"                            DIVISION AFDELING," & _
				"                            TO_NUMBER (TO_CHAR (PLANTINGDATE, 'YYYY')) TAHUN_TANAM," & _
				"                            FF.FCCODE BLOK," & _
				"                            0 QTY" & _
				"                       FROM JOB JJ, FIELD FF" & _
				"                      WHERE (JJ.FCBA = FF.FCBA) AND FF.ACTIVATION = 'Y'" & _
				"                            AND JJ.FCCODE IN" & _
				"                                   ('1320103'," & _
				"                                    '1320104'," & _
				"                                    '1320105'," & _
				"                                    '1320106'," & _
				"                                    '1320401'," & _
				"                                    '1320501'," & _
				"                                    '1320502'," & _
				"                                    '1320508')" & _
				"                     UNION ALL" & _
				"                     SELECT AG.FCBA," & _
				"                            FDDATE TANGGAL," & _
				"                            JOBCODE," & _
				"                            DIVISION AFDELING," & _
				"                            TO_NUMBER (TO_CHAR (PLANTINGDATE, 'YYYY')) TAHUN_TANAM," & _
				"                            LOCATIONCODE BLOK," & _
				"                            OUTPUT QTY" & _
				"                       FROM ATTENDANCE_GAD AG, FIELD FF" & _
				"                      WHERE (AG.FCBA = FF.FCBA AND AG.LOCATIONCODE = FF.FCCODE)" & _
				"                            AND JOBCODE IN" & _
				"                                   ('1320103'," & _
				"                                    '1320104'," & _
				"                                    '1320105'," & _
				"                                    '1320106'," & _
				"                                    '1320401'," & _
				"                                    '1320501'," & _
				"                                    '1320502'," & _
				"                                    '1320508')" & _
				"                     UNION ALL" & _
				"                     SELECT WIP.FCBA," & _
				"                            WIPDATE TANGGAL," & _
				"                            JOB JOBCODE," & _
				"                            DIVISION AFDELING," & _
				"                            TO_NUMBER (TO_CHAR (PLANTINGDATE, 'YYYY')) TAHUN_TANAM," & _
				"                            LOCATIONCODE BLOK," & _
				"                            WIP.QTY" & _
				"                       FROM WORKINPROGRESS WIP, AGREEMENTDETAIL AD, FIELD FF" & _
				"                      WHERE (    WIP.FCBA = AD.FCBA" & _
				"                             AND WIP.AGREEMENTCODE = AD.AGREEMENTCODE" & _
				"                             AND WIP.AGREEMENTLINENO = AD.LINENO)" & _
				"                            AND (AD.FCBA = FF.FCBA AND AD.LOCATIONCODE = FF.FCCODE)" & _
				"                            AND AD.JOB IN" & _
				"                                   ('1320103'," & _
				"                                    '1320104'," & _
				"                                    '1320105'," & _
				"                                    '1320106'," & _
				"                                    '1320401'," & _
				"                                    '1320501'," & _
				"                                    '1320502'," & _
				"                                    '1320508'))" & _
				"              WHERE TO_NUMBER (TO_CHAR (TANGGAL, 'YYYY')) = " & v_txtYear & "" & _
				"                    AND TO_NUMBER (TO_CHAR (TANGGAL, 'MM')) = " & v_txtMonth & "" & _
				"           GROUP BY FCBA," & _
				"                    TANGGAL," & _
				"                    JOBCODE," & _
				"                    AFDELING," & _
				"                    TAHUN_TANAM," & _
				"                    BLOK)" & _
				" GROUP BY FCBA," & _
				"          JOBCODE," & _
				"          AFDELING," & _
				"          TAHUN_TANAM," & _
				"          BLOK"
								
strHarianLC2	=	"     , JOBCODE" & _
					"     , JOBNAME" & _
					"     , SUM(DAY_1) DAY_1" & _
					"     , SUM(DAY_2) DAY_2" & _
					"     , SUM(DAY_3) DAY_3" & _
					"     , SUM(DAY_4) DAY_4" & _
					"     , SUM(DAY_5) DAY_5" & _
					"     , SUM(DAY_6) DAY_6" & _
					"     , SUM(DAY_7) DAY_7" & _
					"     , SUM(DAY_8) DAY_8" & _
					"     , SUM(DAY_9) DAY_9" & _
					"     , SUM(DAY_10) DAY_10" & _
					"     , SUM(DAY_11) DAY_11" & _
					"     , SUM(DAY_12) DAY_12" & _
					"     , SUM(DAY_13) DAY_13" & _
					"     , SUM(DAY_14) DAY_14" & _
					"     , SUM(DAY_15) DAY_15" & _
					"     , SUM(DAY_16) DAY_16" & _
					"     , SUM(DAY_17) DAY_17" & _
					"     , SUM(DAY_18) DAY_18" & _
					"     , SUM(DAY_19) DAY_19" & _
					"     , SUM(DAY_20) DAY_20" & _
					"     , SUM(DAY_21) DAY_21" & _
					"     , SUM(DAY_22) DAY_22" & _
					"     , SUM(DAY_23) DAY_23" & _
					"     , SUM(DAY_24) DAY_24" & _
					"     , SUM(DAY_25) DAY_25" & _
					"     , SUM(DAY_26) DAY_26" & _
					"     , SUM(DAY_27) DAY_27" & _
					"     , SUM(DAY_28) DAY_28" & _
					"     , SUM(DAY_29) DAY_29" & _
					"     , SUM(DAY_30) DAY_30" & _
					"     , SUM(DAY_31) DAY_31" & _
					"     , SUM(REAL) REAL" & _
					"     , SUM(BUDGET) BUDGET"

strHarianLC	=	" select GroupPlantation " & selectCons & strHarianLC2 &"  From (" & strMain & " )" & whereCons & _
					" group by GroupPlantation" & selectCons & " , JOBCODE, JOBNAME" & _
					" having (SUM(REAL) + SUM(BUDGET) > 0)" & _
					" Order by GroupPlantation" & selectCons & ", JOBCODE"

'response.write strHarianLC
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Harian Land Clearing</title>
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
		<div id="header-wrapper">
			<div class="printhide">
				<!--#include file="menutop.asp"-->
			</div>
			<form action="<%=ScriptName%>" method="post" name="frmEntry">
			<div class="printhide">
				<!--#include virtual="/REPORT/FormSelectNewSum.html"-->
				<table id="form-table" class="centered-margin">
					<tr>
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
							<%For i = 1 to 12%>
								<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
							<%Next%>
							</select>
						</td>
						<td>
							<font face="verdana" size="1">JOB CODE</font>
						</td>
						<td>
							<select name="txtJobcode">
								<option value="--ALL--"<%If v_txtJobcode = "--ALL--" then%> Selected <%End If%>>--ALL--</option>
								<%
								Dim rsJobCode, strJobCode
								strJobCode	=	" SELECT DISTINCT FCCODE, FCNAME FROM JOB WHERE FCCODE IN (" & _
												" '1320103', '1320104', '1320105', '1320106'," & _
												" '1320401', '1320501', '1320502', '1320508') ORDER BY FCCODE"
								Set rsJobCode = Server.CreateObject("ADODB.Recordset")
								rsJobCode.Open strJobCode, connstr, adOpenStatic 
								
								if not rsJobCode.eof then
								Do While Not rsJobCode.eof
								if rsJobCode("FCCODE") = v_txtJobcode then%>
								<option value="<%=rsJobCode("FCCODE")%>" Selected><%=rsJobCode("FCCODE")%> : <%=rsJobCode("FCNAME")%></option> 
								<%Else%>
								<option value="<%=rsJobCode("FCCODE")%>"><%=rsJobCode("FCCODE")%> : <%=rsJobCode("FCNAME")%></option>  
								<%End if
									rsJobCode.movenext
								loop
								end if
								rsJobCode.Close
								set rsJobCode = Nothing
								%>
							</select>
						</td>
						<td>
							<font face="verdana" size="1">AFDELING</font>
						</td>
						<td>
							<select name="txtAfdeling" >
								<option value="--SUMMARY--"<%If v_txtAfdeling = "--SUMMARY--" then%> Selected <%End If%>>--SUMMARY--</option>
								<option value="--ALL--"<%If v_txtAfdeling = "--ALL--" then%> Selected <%End If%>>--ALL--</option>
								<%
								Dim rsAfdeling, strAfdeling
								strAfdeling	=	" SELECT DISTINCT FCCODE FROM PARAMETER WHERE PARCODE = 'F0001' ORDER BY FCCODE"
								
								Set rsAfdeling = Server.CreateObject("ADODB.Recordset")
								rsAfdeling.Open strAfdeling, connstr, adOpenStatic 
								
								if not rsAfdeling.eof then
								Do While Not rsAfdeling.eof
								if rsAfdeling("FCCODE") = v_txtAfdeling then%>
								<option value="<%=rsAfdeling("FCCODE")%>" Selected><%=rsAfdeling("FCCODE")%></option> 
								<%Else%>
								<option value="<%=rsAfdeling("FCCODE")%>"><%=rsAfdeling("FCCODE")%></option>  
								<%End if
									rsAfdeling.movenext
								loop
								end if
								rsAfdeling.Close
								set rsAfdeling = Nothing
								%>
							</select>
						</td>
						<td>
							<font face="verdana" size="1">TT</font>
						</td>
						<td>
							<select name="txtTahunTanam" >
								<option value="--SUMMARY--"<%If v_txtTahunTanam = "--SUMMARY--" then%> Selected <%End If%>>--SUMMARY--</option>
								<option value="--ALL--"<%If v_txtTahunTanam = "--ALL--" then%> Selected <%End If%>>--ALL--</option>
								<%
								Dim rsTahunTanam, strTahunTanam
								strTahunTanam	=	" SELECT DISTINCT TO_CHAR(PLANTINGDATE, 'YYYY') TT FROM FIELD WHERE ACTIVATION = 'Y'" & _
													" ORDER BY TO_CHAR(PLANTINGDATE, 'YYYY')"
								
								Set rsTahunTanam = Server.CreateObject("ADODB.Recordset")
								rsTahunTanam.Open strTahunTanam, connstr, adOpenStatic 
								
								if not rsTahunTanam.eof then
								Do While Not rsTahunTanam.eof
								if rsTahunTanam("TT") = v_txtTahunTanam then%>
								<option value="<%=rsTahunTanam("TT")%>" Selected><%=rsTahunTanam("TT")%></option> 
								<%Else%>
								<option value="<%=rsTahunTanam("TT")%>"><%=rsTahunTanam("TT")%></option>  
								<%End if
									rsTahunTanam.movenext
								loop
								end if
								rsTahunTanam.Close
								set rsTahunTanam = Nothing
								%>
							</select>
						</td>
						<td>
							<font face="verdana" size="1">BLOK</font>
						</td>
						<td>
							<select name="txtBlok" >
								<option value="--SUMMARY--"<%If v_txtBlok = "--SUMMARY--" then%> Selected <%End If%>>--SUMMARY--</option>
								<option value="--ALL--"<%If v_txtBlok = "--ALL--" then%> Selected <%End If%>>--ALL--</option>
								<%
								Dim rsBlok, strBlok
								strBlok	=	" SELECT DISTINCT FCCODE FROM FIELD WHERE ACTIVATION = 'Y' ORDER BY FCCODE"
								
								Set rsBlok = Server.CreateObject("ADODB.Recordset")
								rsBlok.Open strBlok, connstr, adOpenStatic 
								
								if not rsBlok.eof then
								Do While Not rsBlok.eof
								if rsBlok("FCCODE") = v_txtBlok then%>
								<option value="<%=rsBlok("FCCODE")%>" Selected><%=rsBlok("FCCODE")%></option> 
								<%Else%>
								<option value="<%=rsBlok("FCCODE")%>"><%=rsBlok("FCCODE")%></option>  
								<%End if
									rsBlok.movenext
								loop
								end if
								rsBlok.Close
								set rsBlok = Nothing
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
            <div id="header">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
            </div>
            <div id="subheader" class="centered">
                <h2>LAPORAN HARIAN LAND CLEARING</h3>
                <br/>
                <h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
            </div>
		</div>


<table border="0" width="100%" id="product-table">
	<tr>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">NAMA PEKEJAAN</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">UOM</th>
		<% if v_txtGroupPlantation <> "--SUMMARY--" and v_txtGroupPlantation <> "" then %>
		    <th rowspan="2" style="background-color:#00A300;color:#fff;">GROUP PLANTATION</th>
		<% else
			nSpanTotal = nSpanTotal - 1
		end if %>
		
		<% if v_txtPlantation <> "--SUMMARY--" and v_txtPlantation <> "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;">PLANTATION</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtGroupRegion <>  "--SUMMARY--" and v_txtGroupRegion <>  "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;">GROUP REGION</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
		      <th rowspan="2" style="background-color:#00A300;color:#fff;">REGION</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
		       <th rowspan="2" style="background-color:#00A300;color:#fff;">COMPANY</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtfcba <>  "--SUMMARY--" and v_txtfcba <>  "" then %>
		       <th rowspan="2" style="background-color:#00A300;color:#fff;">SITE</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>

		<% if v_txtAfdeling <>  "--SUMMARY--" and v_txtAfdeling <>  "" then %>
		       <th rowspan="2" style="background-color:#00A300;color:#fff;">AFDELING</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtTahunTanam <>  "--SUMMARY--" and v_txtTahunTanam <>  "" then %>
		       <th rowspan="2" style="background-color:#00A300;color:#fff;">TAHUN TANAM</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		
		<% if v_txtBlok <>  "--SUMMARY--" and v_txtBlok <>  "" then %>
		       <th rowspan="2" style="background-color:#00A300;color:#fff;">BLOK</th>
		<% else
			nSpanTotal = nSpanTotal - 1
	    end if %>
		<th colspan="31" style="background-color:#00A300;color:#fff;">TANGGAL</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">SHI</th>
	</tr>
	<tr>
		<%for nday = 1 to 31%>
		<th style="background-color:#00A300;color:#fff;"><%=nday%></th>
		<%next%>
		<th style="background-color:#00A300;color:#fff;">REAL</th>
		<th style="background-color:#00A300;color:#fff;">BUDGET</th>
		<th style="background-color:#00A300;color:#fff;">%</th>
	</tr>
<%
Set rsHarianLC = Server.CreateObject("ADODB.Recordset")
rsHarianLC.Open strHarianLC, connstr, adOpenStatic

Dim RowReal : RowReal = 0
Dim RowBudget : RowBudget = 0
Dim RowPercentage : RowPercentage = 0

Dim tempGroupPlantation : tempGroupPlantation = ""
Dim tempPlantation : tempPlantation = ""
Dim tempGroupRegion : tempGroupRegion = ""
Dim tempRegion : tempRegion = ""
Dim tempCompany : tempCompany = ""
Dim tempfcba : tempfcba = ""
Dim tempJobCode : tempJobCode = ""
Dim tempAfdeling : tempAfdeling = ""
Dim tempTahunTanam : tempTahunTanam = ""
Dim tempBlok : tempBlok = ""


Dim SprevGroupPlantation : SprevGroupPlantation = ""
Dim SprevPlantation : SprevPlantation = ""
Dim SprevGroupRegion : SprevGroupRegion = ""
Dim SprevRegion : SprevRegion = ""
Dim SprevCompany : SprevCompany = ""
Dim sprevFcba : sprevFcba = ""
Dim sprevJobCode : sprevJobCode = ""
Dim sprevAfdeling : sprevAfdeling = ""
Dim sprevTahunTanam : sprevTahunTanam = ""
Dim sprevBlok : sprevBlok = ""

if rsHarianLC.bof and rsHarianLC.eof then
%>
	<tr>
		<td colspan="<%=1+nSpanTotal+34%>">NO DATA AVAILABLE</td>
	</tr>
<%
else
		while not rsHarianLC.eof
%>
				<tr>
				<% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <>  "" then %>
					<%
					   tempJobCode = rsHarianLC("JOBCODE")
					   tempGroupPlantation = "--ALL--"
					   tempPlantation = "--SUMMARY--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempAfdeling = "--SUMMARY--"
					   tempTahunTanam = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					<% if tempJobCode <> SprevJobCode then %>
						<td align="left">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
							<%=rsHarianLC("JOBNAME")%>
							</a>
						</td>
						<% SprevJobCode = tempJobCode %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
					<td align="center">HA</td>
				
				<% if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <>  "" then %>
					<% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsHarianLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   tempGroupPlantation = rsHarianLC("GroupPlantation")
					   tempPlantation = "--ALL--"
					   tempGroupRegion = "--SUMMARY--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempAfdeling = "--SUMMARY--"
					   tempTahunTanam = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					<% if tempGroupPlantation <> SprevGroupPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
							<%=rsHarianLC("GroupPlantation")%>
							</a>
						</td>
						<% SprevGroupPlantation = tempGroupPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				
				<% if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <>  "" then %>
					<% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsHarianLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsHarianLC("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if
					   
					   tempPlantation = rsHarianLC("Plantation")
					   tempGroupRegion = "--ALL--"
					   tempRegion = "--SUMMARY--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempAfdeling = "--SUMMARY--"
					   tempTahunTanam = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					<% if tempPlantation <> SprevPlantation then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
							<%=rsHarianLC("Plantation")%>
							</a>
						</td>
						<% SprevPlantation = tempPlantation %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtGroupRegion <>  "--SUMMARY--"  and  v_txtGroupRegion <>  "" then %>
					<% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsHarianLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsHarianLC("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsHarianLC("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   tempGroupRegion = rsHarianLC("GroupRegion")
					   tempRegion = "--ALL--"
					   tempCompany ="--SUMMARY--"
					   tempfcba = "--SUMMARY--"
					   tempAfdeling = "--SUMMARY--"
					   tempTahunTanam = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					<% if tempGroupRegion <> SprevGroupRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
							<%=rsHarianLC("GroupRegion")%>
							</a>
						</td>
						<% SprevGroupRegion= tempGroupRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtRegion <>  "--SUMMARY--" and v_txtRegion <>  "" then %>
				    <% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsHarianLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsHarianLC("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsHarianLC("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--"  and  v_txtGroupRegion <> "" then
					       tempGroupRegion = rsHarianLC("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   tempRegion = rsHarianLC("Region") 
					   tempCompany ="--ALL--"
					   tempfcba = "--SUMMARY--"
					   tempAfdeling = "--SUMMARY--"
					   tempTahunTanam = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					%>
					
					<% if tempRegion <> SprevRegion then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
							<%=rsHarianLC("region")%>
							</a>
						</td>
						<% SprevRegion = tempRegion %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>	
				<% end if %>
				
				<% if v_txtCompany <>  "--SUMMARY--" and v_txtCompany <>  "" then %>
				    <% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsHarianLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsHarianLC("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsHarianLC("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsHarianLC("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsHarianLC("Region") 
					   else
					       tempRegion = ""
					   end if
					   tempCompany =rsHarianLC("Company")
					   tempfcba = "--ALL--"
					   tempAfdeling = "--SUMMARY--"
					   tempTahunTanam = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					 %>
					
					<% if tempCompany <> SprevCompany then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
							<%=rsHarianLC("company")%>
							</a>
						</td>
						<% SprevCompany = tempCompany %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtFcba <>  "--SUMMARY--" and v_txtFcba <>  "" then %>
				    <% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsHarianLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsHarianLC("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsHarianLC("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion  <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsHarianLC("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion  <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsHarianLC("Region") 
					   else
					       tempRegion = ""
					   end if
					   
					   if  v_txtcompany  <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsHarianLC("Company")
					   else
					       tempCompany =""
					   end if
					   
					   tempfcba = rsHarianLC("fcba")
					   tempAfdeling = "--ALL--"
					   tempTahunTanam = "--SUMMARY--"
					   tempBlok = "--SUMMARY--"
					 %>
					
					<% if tempfcba <> sprevFcba then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
						    <%=rsHarianLC("fcba")%>
							</a>
						</td>
						<% sprevFcba = tempfcba %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtAfdeling <>  "--SUMMARY--" and v_txtAfdeling <>  "" then %>
				    <% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsHarianLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsHarianLC("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsHarianLC("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion  <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsHarianLC("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion  <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsHarianLC("Region") 
					   else
					       tempRegion = ""
					   end if
					   
					   if  v_txtcompany  <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsHarianLC("Company")
					   else
					       tempCompany =""
					   end if
					   
					   if  v_txtFcba  <>  "--SUMMARY--" and v_txtFcba <> "" then
					       tempfcba = rsHarianLC("fcba")
					   else
					       tempfcba =""
					   end if

					   tempAfdeling = rsHarianLC("AFDELING")
					   tempTahunTanam = "--ALL--"
					   tempBlok = "--SUMMARY--"
					 %>
					
					<% if tempAfdeling <> sprevAfdeling then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
						    <%=rsHarianLC("AFDELING")%>
							</a>
						</td>
						<% sprevAfdeling = tempAfdeling %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>

				<% if v_txtTahunTanam <>  "--SUMMARY--" and v_txtTahunTanam <>  "" then %>
				    <% if v_txtJobcode <>  "--SUMMARY--"  and  v_txtJobcode <> "" then
					      tempJobCode = rsHarianLC("JobCode")
					   else
					      tempJobCode = ""
					   end if
					   
					   if v_txtGroupPlantation <>  "--SUMMARY--"  and  v_txtGroupPlantation <> "" then
					      tempGroupPlantation = rsHarianLC("GroupPlantation")
					   else
					      tempGroupPlantation = ""
					   end if					
					   
					   if v_txtPlantation <>  "--SUMMARY--"  and  v_txtPlantation <> "" then
					      tempPlantation = rsHarianLC("Plantation")
					   else
					      tempPlantation = ""
					   end if
					   
					   if  v_txtGroupregion  <>  "--SUMMARY--" and v_txtGroupregion <> "" then
					       tempGroupRegion = rsHarianLC("GroupRegion")
					   else
					       tempGroupRegion = ""
					   end if
					   
					   if  v_txtregion  <>  "--SUMMARY--" and v_txtregion <> "" then
					       tempRegion = rsHarianLC("Region") 
					   else
					       tempRegion = ""
					   end if
					   
					   if  v_txtcompany  <>  "--SUMMARY--" and v_txtcompany <> "" then
					       tempCompany =rsHarianLC("Company")
					   else
					       tempCompany =""
					   end if
					   
					   if  v_txtFcba  <>  "--SUMMARY--" and v_txtFcba <> "" then
					       tempfcba = rsHarianLC("fcba")
					   else
					       tempfcba =""
					   end if
					   
					   if  v_txtAfdeling  <>  "--SUMMARY--" and v_txtAfdeling <> "" then
					       tempAfdeling = rsHarianLC("AFDELING")
					   else
					       tempAfdeling =""
					   end if

					   tempTahunTanam = rsHarianLC("TAHUN_TANAM")
					   tempBlok = "--ALL--"
					 %>
					
					<% if tempTahunTanam <> sprevTahunTanam then %>
						<td align="center">
							<a href="<%=ScriptName%>?txtJobcode=<%=tempJobCode%>&txtGroupPlantation=<%=tempGroupPlantation%>&txtplantation=<%=tempPlantation%>&txtgroupregion=<%=tempGroupRegion%>&txtregion=<%=tempRegion%>&txtcompany=<%=tempcompany%>&txtfcba=<%=tempfcba%>&txtyear=<%=v_txtyear%>&txtMonth=<%=v_txtMonth%>&txtAfdeling=<%=tempAfdeling%>&txtTahunTanam=<%=tempTahunTanam%>&txtBlok=<%=tempBlok%>">
						    <%=rsHarianLC("TAHUN_TANAM")%>
							</a>
						</td>
						<% sprevTahunTanam = tempTahunTanam %>	
					<% else	%>
						<td align="center"></td>					
					<% end if %>
				<% end if %>
				
				<% if v_txtBlok <> "--SUMMARY--" and v_txtBlok <>  "" then %>
				    <% if sprevBlok <> rsHarianLC("Blok") then %>				
					    <td align="center"><%=rsHarianLC("Blok")%></td>				
						<% sprevBlok = rsHarianLC("Blok")%>
					<% else	%>
					    <td align="center"></td>					
					<% end if %>							 
				<% end if %>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_1"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_2"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_3"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_4"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_5"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_6"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_7"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_8"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_9"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_10"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_11"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_12"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_13"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_14"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_15"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_16"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_17"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_18"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_19"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_20"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_21"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_22"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_23"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_24"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_25"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_26"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_27"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_28"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_29"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_30"), 0)%></td>
					<td align="right"><%=FormatNumber(rsHarianLC("DAY_31"), 0)%></td>
<%
			RowReal	=	RowReal + Cdbl(rsHarianLC("REAL"))
			RowBudget	=	RowBudget + Cdbl(rsHarianLC("BUDGET"))
			If RowBudget = 0 Then
				RowPercentage = 0
			Else
				RowPercentage = RowReal/RowBudget*100
			End If
%>
					<td align="right"><%=FormatNumber(RowReal, 0)%></td>
					<td align="right"><%=FormatNumber(RowBudget, 0)%></td>
					<td align="right"><%=FormatNumber(RowPercentage, 0)%></td>
				</tr>
<%
			rsHarianLC.MoveNext
			
			RowReal = 0
			RowBudget = 0
			RowPercentage = 0
			Wend	
rsHarianLC.Close
set rsHarianLC = Nothing
%>
<%End if%>
</table>

   </body>
   </html>