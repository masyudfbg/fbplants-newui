<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtFcba, v_txtYear, v_txtMonth, v_txtTanggal, v_txtAfdeling, v_txtTahunTanam, v_txtBlok, LoopNo
Dim ScriptName
Dim rsRotasi, strRotasi, strsql
Dim NO, cClass, nday,ColorID
Dim rsDefaultBa, strDefaultBa, varDefaultBa
Dim rsBaInfo, strBaInfo, varFcbaName, varCompanyName

v_txtYear		=	Request("txtYear")
v_txtMonth		=	Request("txtMonth")
v_txtAfdeling	=	Request("txtAfdeling")
v_txtBlok		=	Request("txtBlok")
v_txtFcba		=	Request("txtFcba")
v_txtTahunTanam	=	Request("txtTahunTanam")

Dim varManager : varManager = ""
Dim varKTU : varKTU = ""

strDefaultBa = " SELECT" & _
				"     FCCODE as FCBA" & _
				" FROM BUSINESSUNIT BU" & _
				" INNER JOIN USER_FCBA_TAB UFT ON (BU.FCCODE = UFT.FCBA) " & _
				" WHERE" & _
				"     ISDEFAULT = 'TRUE'" & _
				"     AND FCTYPE = 'E'" & _
				"     AND FCUSERCODE = '" & Session("CodeName") & "'"

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

If v_txtYear = "" Then
	v_txtYear = Year(Session("v_Costbook"))
End If

If v_txtMonth = "" Then
	v_txtMonth = Month(Session("v_Costbook"))
End If

v_txtTanggal = v_txtMonth & "/" & GetLastDay(v_txtMonth, v_txtYear) & "/" & v_txtYear

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

strRotasi	=	" SELECT" & _
				"     FCBA" & _
				"     , AFDELING" & _
				"     , TAHUNTANAM" & _
				"     , BLOK" & _
				"     , SUM(HEKTAR) HEKTAR" & _
				"     , SUM(DECODE(NMONTH, 1, ROTASI, 0)) ROT_1" & _
				"     , SUM(DECODE(NMONTH, 2, ROTASI, 0)) ROT_2" & _
				"     , SUM(DECODE(NMONTH, 3, ROTASI, 0)) ROT_3" & _
				"     , SUM(DECODE(NMONTH, 4, ROTASI, 0)) ROT_4" & _
				"     , SUM(DECODE(NMONTH, 5, ROTASI, 0)) ROT_5" & _
				"     , SUM(DECODE(NMONTH, 6, ROTASI, 0)) ROT_6" & _
				"     , SUM(DECODE(NMONTH, 7, ROTASI, 0)) ROT_7" & _
				"     , SUM(DECODE(NMONTH, 8, ROTASI, 0)) ROT_8" & _
				"     , SUM(DECODE(NMONTH, 9, ROTASI, 0)) ROT_9" & _
				"     , SUM(DECODE(NMONTH, 10, ROTASI, 0)) ROT_10" & _
				"     , SUM(DECODE(NMONTH, 11, ROTASI, 0)) ROT_11" & _
				"     , SUM(DECODE(NMONTH, 12, ROTASI, 0)) ROT_12" & _
				"     , SUM(DECODE(NYEAR, " & v_txtYear & ", ROTASI, 0))/" & v_txtMonth & " ROT_TOTAL" & _
				"     , SUM(DECODE(NMONTH, 1, TONASE, 0))/1000 TON_1" & _
				"     , SUM(DECODE(NMONTH, 2, TONASE, 0))/1000 TON_2" & _
				"     , SUM(DECODE(NMONTH, 3, TONASE, 0))/1000 TON_3" & _
				"     , SUM(DECODE(NMONTH, 4, TONASE, 0))/1000 TON_4" & _
				"     , SUM(DECODE(NMONTH, 5, TONASE, 0))/1000 TON_5" & _
				"     , SUM(DECODE(NMONTH, 6, TONASE, 0))/1000 TON_6" & _
				"     , SUM(DECODE(NMONTH, 7, TONASE, 0))/1000 TON_7" & _
				"     , SUM(DECODE(NMONTH, 8, TONASE, 0))/1000 TON_8" & _
				"     , SUM(DECODE(NMONTH, 9, TONASE, 0))/1000 TON_9" & _
				"     , SUM(DECODE(NMONTH, 10, TONASE, 0))/1000 TON_10" & _
				"     , SUM(DECODE(NMONTH, 11, TONASE, 0))/1000 TON_11" & _
				"     , SUM(DECODE(NMONTH, 12, TONASE, 0))/1000 TON_12" & _
				"     , SUM(DECODE(NYEAR, " & v_txtYear & ", TONASE, 0))/1000 TON_TOTAL" & _
				"     , SUM(BUDGET)/1000 TON_BUDGET" & _
				" FROM (" & _
				"     SELECT" & _
				"         FCBA" & _
				"         , " & v_txtYear & " NYEAR" & _
				"         , " & v_txtMonth & " NMONTH" & _
				"         , field_api.getValid_Afdeling(FCCODE,Fcba,TO_DATE('" & v_txtTanggal & "','MM/DD/YYYY')) AFDELING" & _
				"         , To_Number(field_api.getValid_plantingyear(FCCODE,Fcba,TO_DATE('" & v_txtTanggal & "','MM/DD/YYYY'))) TAHUNTANAM" & _
				"         , FCCODE BLOK" & _
				"         , HECTARAGEPLANTED HEKTAR" & _
				"         , 0 ROTASI" & _
				"         , 0 TONASE" & _
				"         , NVL(BUDGET_TRANSACTION_PRODQTY_API.GetQuantityUntilMonth(" & _
				"               FCBA" & _
				"               , " & v_txtYear & "" & _
				"               , 1" & _
				"               , 1" & _
				"               , FCCODE" & _
				"               , field_api.getValid_Afdeling(FCCODE,Fcba,TO_DATE('" & v_txtTanggal & "','MM/DD/YYYY'))" & _
				"               , To_Number(field_api.getValid_plantingyear(FCCODE,Fcba,TO_DATE('" & v_txtTanggal & "','MM/DD/YYYY'))) " & _
				"               , TO_DATE('" & v_txtTanggal & "','MM/DD/YYYY')" & _
				"               ), 0) BUDGET" & _
				"     FROM FIELD" & _
				"     WHERE field_api.GetValid_Activation(FCCODE,Fcba,TO_DATE('" & v_txtTanggal & "','MM/DD/YYYY'))= 'Y'" & _
				"     AND field_api.GetValid_status(FCCODE,Fcba,TO_DATE('" & v_txtTanggal & "','MM/DD/YYYY'))= 'TM'" & _
				"     UNION ALL" & _
				"    SELECT FCBA,NYEAR,NMONTH,AFDELING,TAHUNTANAM,BLOK,0 HEKTAR,sum(ROTASI) ROTASI,0 TONASE,0 BUDGET from (   " & _
                "           select HI.FCBA,to_number(to_char(HI.ONDATE,'YYYY')) Nyear,to_number(to_char(HI.ONDATE,'MM')) Nmonth , " & _
                "			field_api.getValid_Afdeling(HI.FIELDCODE,hi.fcba,hi.ondate) Afdeling,to_number(field_api.getValid_plantingyear(HI.FIELDCODE,hi.fcba,hi.ondate)) tahuntanam, " & _
                "			HI.FIELDCODE BLOK,decode(HI.Originalnumber,1,1,0) rotasi " & _
				"			from HARVESTING_INTERVAL_TAB HI ) " & _
				"	 where NYEAR = " & v_txtYear & "" & _
				"	 and NMONTH <=" & v_txtMonth & "" & _
				"	group by FCBA,NYEAR,NMONTH,AFDELING,TAHUNTANAM,BLOK " & _
				"     UNION ALL" & _
				"     SELECT" & _
				"         HV.FCBA" & _
				"         , TO_NUMBER(TO_CHAR(HV.RECEPTIONDATE,'YYYY')) NYEAR" & _
				"         , TO_NUMBER(TO_CHAR(HV.RECEPTIONDATE,'MM')) NMONTH" & _
				"         , field_api.getValid_Afdeling(FF.FCCODE,HV.Fcba,HV.RECEPTIONDATE) AFDELING " & _
                "         , to_number(field_api.getValid_plantingyear(FF.FCCODE,HV.Fcba,HV.RECEPTIONDATE)) TAHUNTANAM " & _
				"         , HV.FIELDCODE BLOK" & _
				"         , 0 HEKTAR" & _
				"         , 0 ROTASI" & _
				"         , HV.MILL_WEIGHT TONASE" & _
				"         , 0 BUDGET" & _
				"     FROM" & _
				"         HARVESTING HV" & _
				"         , FIELD FF" & _
				"     WHERE" & _
				"     (HV.FCBA = FF.FCBA AND HV.FIELDCODE = FF.FCCODE)" & _
				"     AND field_api.GetValid_Activation(FF.FCCODE,HV.Fcba,HV.RECEPTIONDATE) = 'Y'" & _
				"     AND field_api.GetValid_status(FF.FCCODE,HV.Fcba,HV.RECEPTIONDATE)= 'TM'" & _
				"     AND TO_NUMBER(TO_CHAR(HV.RECEPTIONDATE,'YYYY')) = " & v_txtYear & "" & _
				"     AND TO_NUMBER(TO_CHAR(HV.RECEPTIONDATE,'MM')) <= " & v_txtMonth & "" & _
				"     ) WHERE FCBA = '" & v_txtFcba & "'"
				
If Len(Trim(v_txtAfdeling)) > 0 Then
	strRotasi = strRotasi & " AND AFDELING = '" & v_txtAfdeling & "'"
End If

If Len(Trim(v_txtTahunTanam)) > 0 Then
	strRotasi = strRotasi & " AND TAHUNTANAM = '" & v_txtTahunTanam & "'"
End If
	
If Len(Trim(v_txtBlok)) > 0 Then
	strRotasi = strRotasi & " AND BLOK = '" & v_txtBlok & "'"
End If
				
strRotasi	=	strRotasi & " GROUP BY" & _
							"     FCBA" & _
							"     , AFDELING" & _
							"     , TAHUNTANAM" & _
							"     , BLOK"
							
strRotasi	=	strRotasi & " HAVING SUM(DECODE(NYEAR, " & v_txtYear & ", TONASE, 0)) + SUM(BUDGET) > 0"
							
strRotasi	=	strRotasi & " ORDER BY" & _
							"     FCBA" & _
							"     , AFDELING" & _
							"     , TAHUNTANAM" & _
							"     , BLOK"
									
strBaInfo = " SELECT" & _
			"     FCNAME as FCBANAME" & _
			"     , FCCOMPANYNAME as COMPANYNAME" & _
			"     , FCMANAGER" & _
			"     , FCKTU" & _
			" FROM BUSINESSUNIT BU" & _
			" WHERE" & _
			"     FCCODE = '" & v_txtFcba & "'"

Set rsBaInfo = Server.CreateObject("ADODB.Recordset")
rsBaInfo.Open strBaInfo, connstr, adOpenStatic

If Not rsBaInfo.EOF Then
	varFcbaName = rsBaInfo("FCBANAME")
	varCompanyName = rsBaInfo("COMPANYNAME")
	varManager = rsBaInfo("FCMANAGER")
	varKTU = rsBaInfo("FCKTU")
Else
	varFcbaName = ""
	varCompanyName = ""
End If
rsBaInfo.Close
set rsBaInfo = Nothing

'response.write strRotasi
%>

<%
			Dim strTitle, rsTitle
			Dim titleGroupPlantation, titlePlantation, titleGroupRegion, titleRegion, titleCompany, titleSite
			
	        strTitle = " select PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPPLANTATION') titleGP, " & _
							    " PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','PLANTATION') titleP , " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','GROUPREGION') titleGR, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','REGION') titleR, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','COMPANY') titleC, " & _
								" PARAMETERSYSDETAIL_API.GETPARAM1 ('REPORTGROUP','SITE') titleS from dual"
								
			Set rsTitle = Server.CreateObject("ADODB.Recordset")
			'response.write rsTitle
			rsTitle.Open strTitle, connstr, 3,3	

			if not rsTitle.EOF then
				no = 0
				titleGroupPlantation = rsTitle ("titleGP")			
				titlePlantation = rsTitle ("titleP")
				titleGroupRegion = rsTitle ("titleGR")
				titleRegion = rsTitle ("titleR")
				titleCompany = rsTitle ("titleC")
				titleSite = rsTitle ("titleS")
			else
				titleGroupPlantation = "Group Plantation"			
				titlePlantation = "Plantation"
				titleGroupRegion = "Group Region"
				titleRegion = "Region"
				titleCompany = "Company"
				titleSite = "Site"
			end if
			rsTitle.Close
			set rsTitle = Nothing
			
%>
<html>

<head>
<title>ePlaS Reporting - Rotasi Panen</title>
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
							<td><font face="verdana" size="1">TAHUN</font></td>
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
					<%
						For i = MinYr to MaxYr
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(v_txtYear) then%> Selected <%End If%>><%=i%></option>
					<%
						Next
					%>
								</select>
							</td>
							<td><font face="verdana" size="1">BULAN</font></td>
							<td width="2">
								<select name="txtMonth">
					<%
						For i = 1 to 12
					%>
							<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
					<%
						Next
					%>
								</select>
							</td>
							<td>
								<font face="verdana" size="1"><%=titleSite%></font>
							</td>
							<td>
								<select name="txtFcba" >
									<%
									Dim rsFcba, strFcba
									strFcba	=	" SELECT FCCODE, ORDERDATA FROM BUSINESSUNIT" & _
												" WHERE USER_COMPANY_API.GETCOMPANYCODE('" & Session("CodeName") & "',fccode) IS NOT NULL" & _
												" AND FCTYPE = 'E'" & _
												" ORDER BY ORDERDATA"
									
									Set rsFcba = Server.CreateObject("ADODB.Recordset")
									rsFcba.Open strFcba, connstr, adOpenStatic
									
									If rsFcba.Bof Or rsFcba.eof Then
									%>
									<option value=""></option>
									<%
									End If
									
									if not rsFcba.eof then
									Do While Not rsFcba.eof
									%>
									<option value="<%=rsFcba("FCCODE")%>" <%if v_txtFcba = rsFcba("FCCODE") then%> Selected <%End If%>><%=rsFcba("FCCODE")%></option>
									<%
										rsFcba.movenext
									loop
									end if
									rsFcba.Close
									set rsFcba = Nothing
									%>
								</select>
							</td>
							<td>
								<font face="verdana" size="1">AFDELING</font>
							</td>
							<td>
								<select name="txtAfdeling" >
									<option value="">--ALL--</option>
									<%
									Dim rsAfdeling, strAfdeling
									strAfdeling	=	" SELECT FCCODE FROM PARAMETER WHERE PARCODE = 'F0001' AND FCBA = '" & v_txtFcba & "' ORDER BY FCCODE"
									
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
								<font face="verdana" size="1">TAHUN TANAM</font>
							</td>
							<td>
								<select name="txtTahunTanam" >
									<option value="">--ALL--</option>
									<%
									Dim rsTahunTanam, strTahunTanam
									strTahunTanam	=	" SELECT DISTINCT TO_CHAR(PLANTINGDATE,'YYYY') TT FROM FIELD" & _
														" WHERE FCBA = '" & v_txtFcba & "'" & _
														" AND ACTIVATION = 'Y'" & _
														" ORDER BY TO_NUMBER(TO_CHAR(PLANTINGDATE,'YYYY'))"
									
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
									<option value="">--ALL--</option>
									<%
									Dim rsBlok, strBlok
									strBlok	=	" SELECT FCCODE, FCNAME FROM FIELD WHERE FCBA = '" & v_txtFcba & "'" & _
												" AND DX_BLOCK_STATUS = 'TM' ORDER BY FCCODE"
									
									Set rsBlok = Server.CreateObject("ADODB.Recordset")
									rsBlok.Open strBlok, connstr, adOpenStatic 
									
									if not rsBlok.eof then
									Do While Not rsBlok.eof
									if rsBlok("FCCODE") = v_txtBlok then%>
									<option value="<%=rsBlok("FCCODE")%>" Selected><%=rsBlok("FCCODE")%> - <%=rsBlok("FCNAME")%></option> 
									<%Else%>
									<option value="<%=rsBlok("FCCODE")%>"><%=rsBlok("FCCODE")%> - <%=rsBlok("FCNAME")%></option>  
									<%End if
										rsBlok.movenext
									loop
									end if
									rsBlok.Close
									set rsBlok = Nothing
									%>
								</select>
							</td>
							<td><input type="submit" value="View" /></td>
						</tr>				
					</table>
				</div>
			</form>
		</div>

<table border="0" width="100%" id="product-table" class="xthf-site">
	<thead style="border-left:1px #000000 solid;">
	<tr>
         <th colspan="49" style="background:#FFFFFF; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
		 	<div id="header">
				<h4><%If Len(Trim(varCompanyName)) > 0 Then%>PT. <%End If%><%=varCompanyName%></h4>
				<h5><%=varFcbaName%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>ROTASI PANEN</h3>
				<br/>
				<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
			</div>
		 </th>
	</tr>
	<tr>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">AFDELING</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">TT</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">BLOK</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">HA</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">JANUARI</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">FEBRUARI</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">MARET</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">APRIL</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">MEI</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">JUNI</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">JULI</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">AGUSTUS</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">SEPTEMBER</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">OKTOBER</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">NOVEMBER</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">DESEMBER</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">REALISASI</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">BUDGET</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">VARIAN (%)</th>
	</tr>
	<tr>
		<%For LoopNo = 1 to 15%>
		<th style="background-color:#00A300;color:#fff;">ROT</th>
		<th style="background-color:#00A300;color:#fff;">TON</th>
		<th style="background-color:#00A300;color:#fff;">TON /HA</th>
		<%Next%>		
	</tr>
    </thead>
	<tbody>
<%
Set rsRotasi = Server.CreateObject("ADODB.Recordset")
rsRotasi.Open strRotasi, connstr, adOpenStatic

Dim RowRotReal : RowRotReal = 0
Dim RowRotBudget : RowRotBudget = 0
Dim RowTonperHaReal : RowTonperHaReal = 0

Dim RowTonReal : RowTonReal = 0
Dim RowTonBudget : RowTonBudget = 0
Dim RowTonperHaBudget : RowTonperHaBudget = 0

Dim varianceRotasi : varianceRotasi = 0
Dim varianceTonase : varianceTonase = 0
Dim varianceTonperHa : varianceTonperHa = 0

if rsRotasi.eof and rsRotasi.eof then
%>
	<tr>
		<td colspan="49">NO DATA AVAILABLE</td>
	</tr>
<%end if


		while not rsRotasi.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center"><%=rsRotasi("AFDELING")%></td>
				<td align="center"><%=rsRotasi("TAHUNTANAM")%></td>
				<td align="center"><%=rsRotasi("BLOK")%></td>
				<td align="center"><%=rsRotasi("HEKTAR")%></td>
				
				<td align="right"><%=FormatNumber(rsRotasi("ROT_1"), 2)%></td>
				<td align="right"><%=FormatNumber(rsRotasi("TON_1"), 2)%></td>
				<%If Cdbl(rsRotasi("HEKTAR")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsRotasi("TON_1"))/Cdbl(rsRotasi("HEKTAR")), 2)%></td>
				<%End If%>
				
				<td align="right"><%=FormatNumber(rsRotasi("ROT_2"), 2)%></td>
				<td align="right"><%=FormatNumber(rsRotasi("TON_2"), 2)%></td>
				<%If Cdbl(rsRotasi("HEKTAR")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsRotasi("TON_2"))/Cdbl(rsRotasi("HEKTAR")), 2)%></td>
				<%End If%>
				
				<td align="right"><%=FormatNumber(rsRotasi("ROT_3"), 2)%></td>
				<td align="right"><%=FormatNumber(rsRotasi("TON_3"), 2)%></td>
				<%If Cdbl(rsRotasi("HEKTAR")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsRotasi("TON_3"))/Cdbl(rsRotasi("HEKTAR")), 2)%></td>
				<%End If%>
				<td align="right"><%=FormatNumber(rsRotasi("ROT_4"), 2)%></td>
				<td align="right"><%=FormatNumber(rsRotasi("TON_4"), 2)%></td>
				<%If Cdbl(rsRotasi("HEKTAR")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsRotasi("TON_4"))/Cdbl(rsRotasi("HEKTAR")), 2)%></td>
				<%End If%>
				<td align="right"><%=FormatNumber(rsRotasi("ROT_5"), 2)%></td>
				<td align="right"><%=FormatNumber(rsRotasi("TON_5"), 2)%></td>
				<%If Cdbl(rsRotasi("HEKTAR")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsRotasi("TON_5"))/Cdbl(rsRotasi("HEKTAR")), 2)%></td>
				<%End If%>
				<td align="right"><%=FormatNumber(rsRotasi("ROT_6"), 2)%></td>
				<td align="right"><%=FormatNumber(rsRotasi("TON_6"), 2)%></td>
				<%If Cdbl(rsRotasi("HEKTAR")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsRotasi("TON_6"))/Cdbl(rsRotasi("HEKTAR")), 2)%></td>
				<%End If%>
				<td align="right"><%=FormatNumber(rsRotasi("ROT_7"), 2)%></td>
				<td align="right"><%=FormatNumber(rsRotasi("TON_7"), 2)%></td>
				<%If Cdbl(rsRotasi("HEKTAR")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsRotasi("TON_7"))/Cdbl(rsRotasi("HEKTAR")), 2)%></td>
				<%End If%>
				
				<td align="right"><%=FormatNumber(rsRotasi("ROT_8"), 2)%></td>
				<td align="right"><%=FormatNumber(rsRotasi("TON_8"), 2)%></td>
				<%If Cdbl(rsRotasi("HEKTAR")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsRotasi("TON_8"))/Cdbl(rsRotasi("HEKTAR")), 2)%></td>
				<%End If%>
				
				<td align="right"><%=FormatNumber(rsRotasi("ROT_9"), 2)%></td>
				<td align="right"><%=FormatNumber(rsRotasi("TON_9"), 2)%></td>
				<%If Cdbl(rsRotasi("HEKTAR")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsRotasi("TON_9"))/Cdbl(rsRotasi("HEKTAR")), 2)%></td>
				<%End If%>
				
				<td align="right"><%=FormatNumber(rsRotasi("ROT_10"), 2)%></td>
				<td align="right"><%=FormatNumber(rsRotasi("TON_10"), 2)%></td>
				<%If Cdbl(rsRotasi("HEKTAR")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsRotasi("TON_10"))/Cdbl(rsRotasi("HEKTAR")), 2)%></td>
				<%End If%>
				
				<td align="right"><%=FormatNumber(rsRotasi("ROT_11"), 2)%></td>
				<td align="right"><%=FormatNumber(rsRotasi("TON_11"), 2)%></td>
				<%If Cdbl(rsRotasi("HEKTAR")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsRotasi("TON_11"))/Cdbl(rsRotasi("HEKTAR")), 2)%></td>
				<%End If%>
				
				<td align="right"><%=FormatNumber(rsRotasi("ROT_12"), 2)%></td>
				<td align="right"><%=FormatNumber(rsRotasi("TON_12"), 2)%></td>
				<%If Cdbl(rsRotasi("HEKTAR")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsRotasi("TON_12"))/Cdbl(rsRotasi("HEKTAR")), 2)%></td>
				<%End If%>
				
				<td align="right"><%=FormatNumber(rsRotasi("ROT_TOTAL"), 2)%></td>
				<td align="right"><%=FormatNumber(rsRotasi("TON_TOTAL"), 2)%></td>
				<%If Cdbl(rsRotasi("HEKTAR")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsRotasi("TON_TOTAL"))/Cdbl(rsRotasi("HEKTAR")), 2)%></td>
				<%End If%>
				
				<td align="right"><%=FormatNumber(Cint("4"), 2)%></td>
				<td align="right"><%=FormatNumber(rsRotasi("TON_BUDGET"), 2)%></td>
				<%If Cdbl(rsRotasi("HEKTAR")) = 0 Then%>
					<td align="right"><%=FormatNumber(0, 2)%></td>
				<%Else%>
					<td align="right"><%=FormatNumber(Cdbl(rsRotasi("TON_BUDGET"))/Cdbl(rsRotasi("HEKTAR")), 2)%></td>
				<%End If%>
				
				<%
				
				RowRotReal = 0
				RowRotBudget = 0
				RowTonperHaReal = 0

				RowTonReal = 0
				RowTonBudget = 0
				RowTonperHaBudget = 0

				varianceRotasi = 0
				varianceTonase = 0
				varianceTonperHa = 0
			
				RowRotReal = Round(Cdbl(rsRotasi("ROT_TOTAL")), 0)
				RowRotBudget = 4
				
				If RowRotBudget = 0 Then
					varianceRotasi = 0
				Else
					varianceRotasi = (RowRotBudget - RowRotReal)/ RowRotBudget * 100
				End If
				
				RowTonReal = Round(Cdbl(rsRotasi("TON_TOTAL")), 0)
				RowTonBudget = Round(Cdbl(rsRotasi("TON_BUDGET")), 0)
				
				If RowTonBudget = 0 Then
					varianceTonase = 0
				Else
					varianceTonase = (RowTonBudget - RowTonReal)/ RowTonBudget * 100
				End If
				
				If Cdbl(rsRotasi("HEKTAR")) = 0 Then
					RowTonperHaReal = 0
				Else
					RowTonperHaReal = Round(Cdbl(rsRotasi("TON_TOTAL"))/Cdbl(rsRotasi("HEKTAR")), 2)
				End If
				
				If Cdbl(rsRotasi("HEKTAR")) = 0 Then
					RowTonperHaBudget = 0
				Else
					RowTonperHaBudget = Round(Cdbl(rsRotasi("TON_BUDGET"))/Cdbl(rsRotasi("HEKTAR")), 2)
				End If
				
				If RowTonperHaBudget = 0 Then
					varianceTonperHa = 0
				Else
					varianceTonperHa = (RowTonperHaBudget - RowTonperHaReal)/ RowTonperHaBudget * 100
				End If
				
				%>
				<td align="right"><%=FormatNumber(varianceRotasi, 2)%></td>
				<td align="right"><%=FormatNumber(varianceTonase, 2)%></td>
				<td align="right"><%=FormatNumber(varianceTonperHa, 2)%></td>
			</tr>
<%			
			rsRotasi.MoveNext
		Wend
rsRotasi.Close
set rsRotasi = Nothing
%>
</tbody>
</table>

   <table width="90%" class="data">
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
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	   </tr>
	   <tr> 
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=varManager%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=varKTU%></u></strong></td>
	   </tr>
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>
   </body>
   </html>
