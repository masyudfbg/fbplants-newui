<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
	Dim v_txtYear, v_txtMonth, v_txtTanggal, v_txtAfdeling, v_txtBlok
	Dim ScriptName
	Dim rsPlantedOutput, strPlantedOutput
	Dim v_txtJobCode
	Dim NO, cClass, val_loop
	
	Dim Rotasi_1, Rotasi_2, Rotasi_3, Rotasi_4, Rotasi_5, Rotasi_6, Rotasi_7, Rotasi_8
	Dim Rotasi_9, Rotasi_10, Rotasi_11, Rotasi_12, Rotasi_13, Rotasi_14, Rotasi_15, Rotasi_16
	Dim Rotasi_17, Rotasi_18, Rotasi_19, Rotasi_20, Rotasi_21, Rotasi_22, Rotasi_23, Rotasi_24
	Dim Rotasi_25, Rotasi_26, Rotasi_27, Rotasi_28, Rotasi_29, Rotasi_30, Rotasi_31, Rotasi_Total
	
	Dim GrandPlanted : GrandPlanted = 0
	
	Dim GrandOutput_1 : GrandOutput_1 = 0
	Dim GrandOutput_2 : GrandOutput_2 = 0
	Dim GrandOutput_3 : GrandOutput_3 = 0
	Dim GrandOutput_4 : GrandOutput_4 = 0
	Dim GrandOutput_5 : GrandOutput_5 = 0
	Dim GrandOutput_6 : GrandOutput_6 = 0
	Dim GrandOutput_7 : GrandOutput_7 = 0
	Dim GrandOutput_8 : GrandOutput_8 = 0
	Dim GrandOutput_9 : GrandOutput_9 = 0
	Dim GrandOutput_10 : GrandOutput_10 = 0
	Dim GrandOutput_11 : GrandOutput_11 = 0
	Dim GrandOutput_12 : GrandOutput_12 = 0
	Dim GrandOutput_13 : GrandOutput_13 = 0
	Dim GrandOutput_14 : GrandOutput_14 = 0
	Dim GrandOutput_15 : GrandOutput_15 = 0
	Dim GrandOutput_16 : GrandOutput_16 = 0
	Dim GrandOutput_17 : GrandOutput_17 = 0
	Dim GrandOutput_18 : GrandOutput_18 = 0
	Dim GrandOutput_19 : GrandOutput_19 = 0
	Dim GrandOutput_20 : GrandOutput_20 = 0
	Dim GrandOutput_21 : GrandOutput_21 = 0
	Dim GrandOutput_22 : GrandOutput_22 = 0
	Dim GrandOutput_23 : GrandOutput_23 = 0
	Dim GrandOutput_24 : GrandOutput_24 = 0
	Dim GrandOutput_25 : GrandOutput_25 = 0
	Dim GrandOutput_26 : GrandOutput_26 = 0
	Dim GrandOutput_27 : GrandOutput_27 = 0
	Dim GrandOutput_28 : GrandOutput_28 = 0
	Dim GrandOutput_29 : GrandOutput_29 = 0
	Dim GrandOutput_30 : GrandOutput_30 = 0
	Dim GrandOutput_31 : GrandOutput_31 = 0
	Dim GrandOutput_Total : GrandOutput_Total = 0
	
	Dim GrandRotasi_1 : GrandRotasi_1 = 0
	Dim GrandRotasi_2 : GrandRotasi_2 = 0
	Dim GrandRotasi_3 : GrandRotasi_3 = 0
	Dim GrandRotasi_4 : GrandRotasi_4 = 0
	Dim GrandRotasi_5 : GrandRotasi_5 = 0
	Dim GrandRotasi_6 : GrandRotasi_6 = 0
	Dim GrandRotasi_7 : GrandRotasi_7 = 0
	Dim GrandRotasi_8 : GrandRotasi_8 = 0
	Dim GrandRotasi_9 : GrandRotasi_9 = 0
	Dim GrandRotasi_10 : GrandRotasi_10 = 0
	Dim GrandRotasi_11 : GrandRotasi_11 = 0
	Dim GrandRotasi_12 : GrandRotasi_12 = 0
	Dim GrandRotasi_13 : GrandRotasi_13 = 0
	Dim GrandRotasi_14 : GrandRotasi_14 = 0
	Dim GrandRotasi_15 : GrandRotasi_15 = 0
	Dim GrandRotasi_16 : GrandRotasi_16 = 0
	Dim GrandRotasi_17 : GrandRotasi_17 = 0
	Dim GrandRotasi_18 : GrandRotasi_18 = 0
	Dim GrandRotasi_19 : GrandRotasi_19 = 0
	Dim GrandRotasi_20 : GrandRotasi_20 = 0
	Dim GrandRotasi_21 : GrandRotasi_21 = 0
	Dim GrandRotasi_22 : GrandRotasi_22 = 0
	Dim GrandRotasi_23 : GrandRotasi_23 = 0
	Dim GrandRotasi_24 : GrandRotasi_24 = 0
	Dim GrandRotasi_25 : GrandRotasi_25 = 0
	Dim GrandRotasi_26 : GrandRotasi_26 = 0
	Dim GrandRotasi_27 : GrandRotasi_27 = 0
	Dim GrandRotasi_28 : GrandRotasi_28 = 0
	Dim GrandRotasi_29 : GrandRotasi_29 = 0
	Dim GrandRotasi_30 : GrandRotasi_30 = 0
	Dim GrandRotasi_31 : GrandRotasi_31 = 0
	Dim GrandRotasi_Total : GrandRotasi_Total = 0
	
	ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

	v_txtYear		=	Request("txtYear")
	v_txtMonth		=	Request("txtMonth")
	v_txtJobCode	=	Request("txtJobCode")
	v_txtAfdeling	=	Request("txtAfdeling")
	v_txtBlok		=	Request("txtBlok")

	If Len(Trim(v_txtYear)) = 0 Then
		v_txtYear = Year(Session("v_Costbook"))
	End If

	If Len(Trim(v_txtMonth)) = 0 Then
		v_txtMonth = Month(Session("v_Costbook"))
	End If
	
	
	v_txtTanggal = v_txtMonth & "/1/" & v_txtYear
	
	

	If Len(Trim(v_txtJobCode)) = 0 Then
		v_txtJobCode = "6502101"
	End If	
	
	strPlantedOutput		=	" SELECT" & _
								"     FCBA" & _
								"     , JOBCODE" & _
								"     , AFDELING" & _
								"     , BLOK" & _
								"     , SUM(PLANTED) PLANTED" & _
								"     , SUM(OUTPUT_1) OUTPUT_1" & _
								"     , SUM(OUTPUT_2) OUTPUT_2" & _
								"     , SUM(OUTPUT_3) OUTPUT_3" & _
								"     , SUM(OUTPUT_4) OUTPUT_4" & _
								"     , SUM(OUTPUT_5) OUTPUT_5" & _
								"     , SUM(OUTPUT_6) OUTPUT_6" & _
								"     , SUM(OUTPUT_7) OUTPUT_7" & _
								"     , SUM(OUTPUT_8) OUTPUT_8" & _
								"     , SUM(OUTPUT_9) OUTPUT_9" & _
								"     , SUM(OUTPUT_10) OUTPUT_10" & _
								"     , SUM(OUTPUT_11) OUTPUT_11" & _
								"     , SUM(OUTPUT_12) OUTPUT_12" & _
								"     , SUM(OUTPUT_13) OUTPUT_13" & _
								"     , SUM(OUTPUT_14) OUTPUT_14" & _
								"     , SUM(OUTPUT_15) OUTPUT_15" & _
								"     , SUM(OUTPUT_16) OUTPUT_16" & _
								"     , SUM(OUTPUT_17) OUTPUT_17" & _
								"     , SUM(OUTPUT_18) OUTPUT_18" & _
								"     , SUM(OUTPUT_19) OUTPUT_19" & _
								"     , SUM(OUTPUT_20) OUTPUT_20" & _
								"     , SUM(OUTPUT_21) OUTPUT_21" & _
								"     , SUM(OUTPUT_22) OUTPUT_22" & _
								"     , SUM(OUTPUT_23) OUTPUT_23" & _
								"     , SUM(OUTPUT_24) OUTPUT_24" & _
								"     , SUM(OUTPUT_25) OUTPUT_25" & _
								"     , SUM(OUTPUT_26) OUTPUT_26" & _
								"     , SUM(OUTPUT_27) OUTPUT_27" & _
								"     , SUM(OUTPUT_28) OUTPUT_28" & _
								"     , SUM(OUTPUT_29) OUTPUT_29" & _
								"     , SUM(OUTPUT_30) OUTPUT_30" & _
								"     , SUM(OUTPUT_31) OUTPUT_31" & _
								"     , SUM(OUTPUT_TOTAL) OUTPUT_TOTAL" & _
								" FROM (" & _
								"     SELECT" & _
								"         FCBA" & _
								"         , JOBCODE" & _
								"         , AFDELING" & _
								"         , BLOK" & _
								"         , FIELD_API.GET_HA(BLOK, FCBA) PLANTED" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 1, OUTPUT, 0)) OUTPUT_1" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 2, OUTPUT, 0)) OUTPUT_2" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 3, OUTPUT, 0)) OUTPUT_3" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 4, OUTPUT, 0)) OUTPUT_4" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 5, OUTPUT, 0)) OUTPUT_5" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 6, OUTPUT, 0)) OUTPUT_6" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 7, OUTPUT, 0)) OUTPUT_7" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 8, OUTPUT, 0)) OUTPUT_8" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 9, OUTPUT, 0)) OUTPUT_9" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 10, OUTPUT, 0)) OUTPUT_10" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 11, OUTPUT, 0)) OUTPUT_11" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 12, OUTPUT, 0)) OUTPUT_12" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 13, OUTPUT, 0)) OUTPUT_13" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 14, OUTPUT, 0)) OUTPUT_14" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 15, OUTPUT, 0)) OUTPUT_15" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 16, OUTPUT, 0)) OUTPUT_16" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 17, OUTPUT, 0)) OUTPUT_17" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 18, OUTPUT, 0)) OUTPUT_18" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 19, OUTPUT, 0)) OUTPUT_19" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 20, OUTPUT, 0)) OUTPUT_20" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 21, OUTPUT, 0)) OUTPUT_21" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 22, OUTPUT, 0)) OUTPUT_22" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 23, OUTPUT, 0)) OUTPUT_23" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 24, OUTPUT, 0)) OUTPUT_24" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 25, OUTPUT, 0)) OUTPUT_25" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 26, OUTPUT, 0)) OUTPUT_26" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 27, OUTPUT, 0)) OUTPUT_27" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 28, OUTPUT, 0)) OUTPUT_28" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 29, OUTPUT, 0)) OUTPUT_29" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 30, OUTPUT, 0)) OUTPUT_30" & _
								"         , SUM(DECODE(TO_NUMBER(TO_CHAR(TANGGAL, 'DD')), 31, OUTPUT, 0)) OUTPUT_31" & _
								"         , SUM(" & _
								"             DECODE(" & _
								"                 TO_NUMBER(TO_CHAR(TANGGAL, 'YYYY')), " & v_txtYear & ", OUTPUT," & _
								"                 TO_NUMBER(TO_CHAR(TANGGAL, 'MM')), " & v_txtMonth & ", OUTPUT, 0)" & _
								"             ) OUTPUT_TOTAL" & _
								"     FROM (" & _
								"         SELECT" & _
								"             FCBA" & _
								"             , TANGGAL" & _
								"             , JOBCODE" & _
								"             , FIELD_API.GET_AFDELING(BLOK, FCBA) AFDELING" & _
								"             , BLOK" & _
								"             , SUM(OUTPUT) OUTPUT" & _
								"         FROM (" & _
								"             SELECT" & _
								"                 FCBA" & _
								"                 , FDDATE TANGGAL" & _
								"                 , JOBCODE" & _
								"                 , LOCATIONCODE BLOK" & _
								"                 , NVL(OUTPUT, 0)  OUTPUT" & _
								"             FROM ATTENDANCE_GAD" & _
								"             WHERE LOCATIONTYPE = 'FF'" & _
								"             AND TO_NUMBER(TO_CHAR(FDDATE, 'YYYY')) = " & v_txtYear & "" & _
								"             AND TO_NUMBER(TO_CHAR(FDDATE, 'MM')) = " & v_txtMonth & "" & _
								"             AND OUTPUT > 0" & _
								"             )" & _
								"         GROUP BY" & _
								"             FCBA" & _
								"             , TANGGAL" & _
								"             , JOBCODE" & _
								"             , BLOK" & _
								"         ) WHERE FCBA = '"& Session("FCBA") & "'" & _
								"                 AND JOBCODE = '" & v_txtJobCode & "'"

If Len(Trim(v_txtAfdeling)) > 0 Then
	strPlantedOutput	=	strPlantedOutput & " AND AFDELING = '" & v_txtAfdeling & "'"
End If

If Len(Trim(v_txtBlok)) > 0 Then
	strPlantedOutput	=	strPlantedOutput & " AND BLOK = '" & v_txtBlok & "'"
End If

strPlantedOutput	=	strPlantedOutput &	" GROUP BY" & _
											"     FCBA" & _
											"     , JOBCODE" & _
											"     , AFDELING" & _
											"     , BLOK" & _
											" )" & _
											" GROUP BY" & _
											"     FCBA" & _
											"     , JOBCODE" & _
											"     , AFDELING" & _
											"     , BLOK"
											
strPlantedOutput	=	strPlantedOutput &	" ORDER BY BLOK"

	'Response.Write strPlantedOutput
	
	Set rsPlantedOutput = Server.CreateObject("ADODB.Recordset")
	rsPlantedOutput.Open strPlantedOutput, connstr, adOpenStatic 
	
%>


<html>

<head>
<title>ePlaS Reporting - Crosscek Ha Planted vs Ha Output</title>
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
						<td>
							<font face="verdana" size="1">TAHUN</font>
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
							<font face="verdana" size="1">AFDELING</font>
						</td>
						<td>
							<select name="txtAfdeling" >
								<option value = "" <%If Len(Trim(v_txtAfdeling)) = 0 then%>  Selected <%End If%>>--ALL--</option>
								<%
								Dim rsAfdeling, strAfdeling
								strAfdeling	=	" SELECT FCCODE FROM PARAMETER WHERE PARCODE = 'F0001' AND FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE"
								
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
							<font face="verdana" size="1">BLOK</font>
						</td>
						<td>
							<select name="txtBlok" >
								<option value = "" <%If Len(Trim(v_txtBlok)) = 0 then%> Selected <%End If%>>--ALL--</option>
								<%
								Dim rsBlok, strBlok
								strBlok	=	" SELECT FCCODE, FCNAME FROM FIELD WHERE FCBA = '" & Session("FCBA") & "'" & _
											" ORDER BY FCCODE"
								
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
						<td>
							<font face="verdana" size="1">JOBCODE</font>
						</td>
						<td>
							<select name="txtJobCode" >
								<%
								Dim rsJobCode, strJobCode
								strJobCode	=	" SELECT DISTINCT JOBCODE FCCODE, JOB_API.get_jobcode_name(JOBCODE, FCBA) FCNAME FROM ATTENDANCE_GAD" & _
												" WHERE FCBA = '" & Session("FCBA") & "'" & _
												"       AND LOCATIONTYPE = 'FF'" & _
												" ORDER BY JOBCODE"
								
								Set rsJobCode = Server.CreateObject("ADODB.Recordset")
								rsJobCode.Open strJobCode, connstr, adOpenStatic 
								
								if not rsJobCode.eof then
								Do While Not rsJobCode.eof
								if rsJobCode("FCCODE") = v_txtJobCode then%>
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
    <th colspan="50" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
				<h4>PT. <%=Session("COMPANY")%></h4>
				<h5><%=Session("BA")%></h5>
			</div>
			
			<div id="subheader" class="centered" tyle="width:50%; margin:auto;">
				<h2>CROSSCEK HA PLANTED VS HA OUTPUT</h2>
                <br/>
		<h4>PERIODE&nbsp;:&nbsp;<%=textMonth(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
			</div>
		</div>
</th>
		</tr>
	<tr>
			<th rowspan="4" style="background-color:#00A300;color:#fff;">BLOK</th>
			<th colspan="96" style="background-color:#00A300;color:#fff;">TANGGAL</th>
		</tr>
		
		<tr>
			<%For val_loop = 1 To 31%>
				<th colspan="3" style="background-color:#00A300;color:#fff;"><%=val_loop%></th>
			<%Next%>
				<th colspan="3" style="background-color:#00A300;color:#fff;">TOTAL</th>
		</tr>
		<tr>
			<%For val_loop = 1 To 32%>
				<th colspan="2" style="background-color:#00A300;color:#fff;">HA</th>
				<th rowspan="2" style="background-color:#00A300;color:#fff;">ROTASI</th>
			<%Next%>
		</tr>
		<tr>
			<%For val_loop = 1 To 32%>
				<th style="background-color:#00A300;color:#fff;">PLANTED</th>
				<th style="background-color:#00A300;color:#fff;">OUTPUT</th>
			<%Next%>
		</tr>
    </thead>
	<tbody>    
<%if rsPlantedOutput.BOF and rsPlantedOutput.EOF then%>
		<tr>
			<td colspan="97" align="left">NO DATA AVAILABLE</td>
		</tr>
<%Else
	Rotasi_1 = 0
	Rotasi_2 = 0
	Rotasi_3 = 0
	Rotasi_4 = 0
	Rotasi_5 = 0
	Rotasi_6 = 0
	Rotasi_7 = 0
	Rotasi_8 = 0
	Rotasi_9 = 0
	Rotasi_10 = 0
	Rotasi_11 = 0
	Rotasi_12 = 0
	Rotasi_13 = 0
	Rotasi_14 = 0
	Rotasi_15 = 0
	Rotasi_16 = 0
	Rotasi_17 = 0
	Rotasi_18 = 0
	Rotasi_19 = 0
	Rotasi_20 = 0
	Rotasi_21 = 0
	Rotasi_22 = 0
	Rotasi_23 = 0
	Rotasi_24 = 0
	Rotasi_25 = 0
	Rotasi_26 = 0
	Rotasi_27 = 0
	Rotasi_28 = 0
	Rotasi_29 = 0
	Rotasi_30 = 0
	Rotasi_31 = 0
	Rotasi_Total = 0

		while not rsPlantedOutput.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
			
			If Cdbl(rsPlantedOutput("PLANTED")) = 0 Then
				Rotasi_1 = 0
				Rotasi_2 = 0
				Rotasi_3 = 0
				Rotasi_4 = 0
				Rotasi_5 = 0
				Rotasi_6 = 0
				Rotasi_7 = 0
				Rotasi_8 = 0
				Rotasi_9 = 0
				Rotasi_10 = 0
				Rotasi_11 = 0
				Rotasi_12 = 0
				Rotasi_13 = 0
				Rotasi_14 = 0
				Rotasi_15 = 0
				Rotasi_16 = 0
				Rotasi_17 = 0
				Rotasi_18 = 0
				Rotasi_19 = 0
				Rotasi_20 = 0
				Rotasi_21 = 0
				Rotasi_22 = 0
				Rotasi_23 = 0
				Rotasi_24 = 0
				Rotasi_25 = 0
				Rotasi_26 = 0
				Rotasi_27 = 0
				Rotasi_28 = 0
				Rotasi_29 = 0
				Rotasi_30 = 0
				Rotasi_31 = 0
				Rotasi_Total = 0
			Else
				Rotasi_1 = Cdbl(rsPlantedOutput("OUTPUT_1"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_2 = Cdbl(rsPlantedOutput("OUTPUT_2"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_3 = Cdbl(rsPlantedOutput("OUTPUT_3"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_4 = Cdbl(rsPlantedOutput("OUTPUT_4"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_5 = Cdbl(rsPlantedOutput("OUTPUT_5"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_6 = Cdbl(rsPlantedOutput("OUTPUT_6"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_7 = Cdbl(rsPlantedOutput("OUTPUT_7"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_8 = Cdbl(rsPlantedOutput("OUTPUT_8"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_9 = Cdbl(rsPlantedOutput("OUTPUT_9"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_10 = Cdbl(rsPlantedOutput("OUTPUT_10"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_11 = Cdbl(rsPlantedOutput("OUTPUT_11"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_12 = Cdbl(rsPlantedOutput("OUTPUT_12"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_13 = Cdbl(rsPlantedOutput("OUTPUT_13"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_14 = Cdbl(rsPlantedOutput("OUTPUT_14"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_15 = Cdbl(rsPlantedOutput("OUTPUT_15"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_16 = Cdbl(rsPlantedOutput("OUTPUT_16"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_17 = Cdbl(rsPlantedOutput("OUTPUT_17"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_18 = Cdbl(rsPlantedOutput("OUTPUT_18"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_19 = Cdbl(rsPlantedOutput("OUTPUT_19"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_20 = Cdbl(rsPlantedOutput("OUTPUT_20"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_21 = Cdbl(rsPlantedOutput("OUTPUT_21"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_22 = Cdbl(rsPlantedOutput("OUTPUT_22"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_23 = Cdbl(rsPlantedOutput("OUTPUT_23"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_24 = Cdbl(rsPlantedOutput("OUTPUT_24"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_25 = Cdbl(rsPlantedOutput("OUTPUT_25"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_26 = Cdbl(rsPlantedOutput("OUTPUT_26"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_27 = Cdbl(rsPlantedOutput("OUTPUT_27"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_28 = Cdbl(rsPlantedOutput("OUTPUT_28"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_29 = Cdbl(rsPlantedOutput("OUTPUT_29"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_30 = Cdbl(rsPlantedOutput("OUTPUT_30"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_31 = Cdbl(rsPlantedOutput("OUTPUT_31"))/Cdbl(rsPlantedOutput("PLANTED"))
				Rotasi_Total = Cdbl(rsPlantedOutput("OUTPUT_TOTAL"))/Cdbl(rsPlantedOutput("PLANTED"))
			End If
%>
		<tr class="<%=cClass%>">
			<td align="left"><%=rsPlantedOutput("BLOK")%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_1"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_1, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_2"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_2, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_3"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_3, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_4"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_4, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_5"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_5, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_6"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_6, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_7"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_7, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_8"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_8, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_9"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_9, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_10"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_10, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_11"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_11, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_12"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_12, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_13"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_13, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_14"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_14, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_15"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_15, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_16"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_16, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_17"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_17, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_18"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_18, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_19"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_19, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_20"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_20, 0)%></td>			
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_21"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_21, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_22"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_22, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_23"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_23, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_24"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_24, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_25"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_25, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_26"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_26, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_27"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_27, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_28"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_28, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_29"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_29, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_30"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_30, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_31"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_31, 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("PLANTED"), 0)%></td>
			<td align="right"><%=FormatNumber(rsPlantedOutput("OUTPUT_TOTAL"), 0)%></td>
			<td align="right"><%=FormatNumber(Rotasi_Total, 1)%></td>
		</tr>
<%
			GrandPlanted	=	GrandPlanted + CdbL(rsPlantedOutput("PLANTED"))
			
			GrandOutput_1 = GrandOutput_1 + Cdbl(rsPlantedOutput("OUTPUT_1"))
			GrandOutput_2 = GrandOutput_2 + Cdbl(rsPlantedOutput("OUTPUT_2"))
			GrandOutput_3 = GrandOutput_3 + Cdbl(rsPlantedOutput("OUTPUT_3"))
			GrandOutput_4 = GrandOutput_4 + Cdbl(rsPlantedOutput("OUTPUT_4"))
			GrandOutput_5 = GrandOutput_5 + Cdbl(rsPlantedOutput("OUTPUT_5"))
			GrandOutput_6 = GrandOutput_6 + Cdbl(rsPlantedOutput("OUTPUT_6"))
			GrandOutput_7 = GrandOutput_7 + Cdbl(rsPlantedOutput("OUTPUT_7"))
			GrandOutput_8 = GrandOutput_8 + Cdbl(rsPlantedOutput("OUTPUT_8"))
			GrandOutput_9 = GrandOutput_9 + Cdbl(rsPlantedOutput("OUTPUT_9"))
			GrandOutput_10 = GrandOutput_10 + Cdbl(rsPlantedOutput("OUTPUT_10"))
			GrandOutput_11 = GrandOutput_11 + Cdbl(rsPlantedOutput("OUTPUT_11"))
			GrandOutput_12 = GrandOutput_12 + Cdbl(rsPlantedOutput("OUTPUT_12"))
			GrandOutput_13 = GrandOutput_13 + Cdbl(rsPlantedOutput("OUTPUT_13"))
			GrandOutput_14 = GrandOutput_14 + Cdbl(rsPlantedOutput("OUTPUT_14"))
			GrandOutput_15 = GrandOutput_15 + Cdbl(rsPlantedOutput("OUTPUT_15"))
			GrandOutput_16 = GrandOutput_16 + Cdbl(rsPlantedOutput("OUTPUT_16"))
			GrandOutput_17 = GrandOutput_17 + Cdbl(rsPlantedOutput("OUTPUT_17"))
			GrandOutput_18 = GrandOutput_18 + Cdbl(rsPlantedOutput("OUTPUT_18"))
			GrandOutput_19 = GrandOutput_19 + Cdbl(rsPlantedOutput("OUTPUT_19"))
			GrandOutput_20 = GrandOutput_20 + Cdbl(rsPlantedOutput("OUTPUT_20"))
			GrandOutput_21 = GrandOutput_21 + Cdbl(rsPlantedOutput("OUTPUT_21"))
			GrandOutput_22 = GrandOutput_22 + Cdbl(rsPlantedOutput("OUTPUT_22"))
			GrandOutput_23 = GrandOutput_23 + Cdbl(rsPlantedOutput("OUTPUT_23"))
			GrandOutput_24 = GrandOutput_24 + Cdbl(rsPlantedOutput("OUTPUT_24"))
			GrandOutput_25 = GrandOutput_25 + Cdbl(rsPlantedOutput("OUTPUT_25"))
			GrandOutput_26 = GrandOutput_26 + Cdbl(rsPlantedOutput("OUTPUT_26"))
			GrandOutput_27 = GrandOutput_27 + Cdbl(rsPlantedOutput("OUTPUT_27"))
			GrandOutput_28 = GrandOutput_28 + Cdbl(rsPlantedOutput("OUTPUT_28"))
			GrandOutput_29 = GrandOutput_29 + Cdbl(rsPlantedOutput("OUTPUT_29"))
			GrandOutput_30 = GrandOutput_30 + Cdbl(rsPlantedOutput("OUTPUT_30"))
			GrandOutput_31 = GrandOutput_31 + Cdbl(rsPlantedOutput("OUTPUT_31"))
			GrandOutput_Total = GrandOutput_Total + Cdbl(rsPlantedOutput("OUTPUT_TOTAL"))
			
			rsPlantedOutput.MoveNext
		Wend
	rsPlantedOutput.Close
	set rsPlantedOutput = Nothing
%>

<%
			If GrandPlanted = 0 Then
				GrandRotasi_1 = 0
				GrandRotasi_2 = 0
				GrandRotasi_3 = 0
				GrandRotasi_4 = 0
				GrandRotasi_5 = 0
				GrandRotasi_6 = 0
				GrandRotasi_7 = 0
				GrandRotasi_8 = 0
				GrandRotasi_9 = 0
				GrandRotasi_10 = 0
				GrandRotasi_11 = 0
				GrandRotasi_12 = 0
				GrandRotasi_13 = 0
				GrandRotasi_14 = 0
				GrandRotasi_15 = 0
				GrandRotasi_16 = 0
				GrandRotasi_17 = 0
				GrandRotasi_18 = 0
				GrandRotasi_19 = 0
				GrandRotasi_20 = 0
				GrandRotasi_21 = 0
				GrandRotasi_22 = 0
				GrandRotasi_23 = 0
				GrandRotasi_24 = 0
				GrandRotasi_25 = 0
				GrandRotasi_26 = 0
				GrandRotasi_27 = 0
				GrandRotasi_28 = 0
				GrandRotasi_29 = 0
				GrandRotasi_30 = 0
				GrandRotasi_31 = 0
				GrandRotasi_Total = 0
			Else
				GrandRotasi_1 = GrandOutput_1/GrandPlanted
				GrandRotasi_2 = GrandOutput_2/GrandPlanted
				GrandRotasi_3 = GrandOutput_3/GrandPlanted
				GrandRotasi_4 = GrandOutput_4/GrandPlanted
				GrandRotasi_5 = GrandOutput_5/GrandPlanted
				GrandRotasi_6 = GrandOutput_6/GrandPlanted
				GrandRotasi_7 = GrandOutput_7/GrandPlanted
				GrandRotasi_8 = GrandOutput_8/GrandPlanted
				GrandRotasi_9 = GrandOutput_9/GrandPlanted
				GrandRotasi_10 = GrandOutput_10/GrandPlanted
				GrandRotasi_11 = GrandOutput_11/GrandPlanted
				GrandRotasi_12 = GrandOutput_12/GrandPlanted
				GrandRotasi_13 = GrandOutput_13/GrandPlanted
				GrandRotasi_14 = GrandOutput_14/GrandPlanted
				GrandRotasi_15 = GrandOutput_15/GrandPlanted
				GrandRotasi_16 = GrandOutput_16/GrandPlanted
				GrandRotasi_17 = GrandOutput_17/GrandPlanted
				GrandRotasi_18 = GrandOutput_18/GrandPlanted
				GrandRotasi_19 = GrandOutput_19/GrandPlanted
				GrandRotasi_20 = GrandOutput_20/GrandPlanted
				GrandRotasi_21 = GrandOutput_21/GrandPlanted
				GrandRotasi_22 = GrandOutput_22/GrandPlanted
				GrandRotasi_23 = GrandOutput_23/GrandPlanted
				GrandRotasi_24 = GrandOutput_24/GrandPlanted
				GrandRotasi_25 = GrandOutput_25/GrandPlanted
				GrandRotasi_26 = GrandOutput_26/GrandPlanted
				GrandRotasi_27 = GrandOutput_27/GrandPlanted
				GrandRotasi_28 = GrandOutput_28/GrandPlanted
				GrandRotasi_29 = GrandOutput_29/GrandPlanted
				GrandRotasi_30 = GrandOutput_30/GrandPlanted
				GrandRotasi_31 = GrandOutput_31/GrandPlanted
				GrandRotasi_Total = GrandOutput_TOTAL/GrandPlanted
			End If
%>
		<tr bgcolor="#F4F4A0">
			<td align="Center"><strong><i>TOTAL</i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_1, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_1, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_2, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_2, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_3, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_3, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_4, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_4, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_5, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_5, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_6, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_6, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_7, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_7, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_8, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_8, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_9, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_9, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_10, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_10, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_11, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_11, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_12, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_12, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_13, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_13, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_14, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_14, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_15, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_15, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_16, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_16, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_17, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_17, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_18, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_18, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_19, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_19, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_20, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_20, 0)%></i></strong></td>			
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_21, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_21, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_22, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_22, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_23, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_23, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_24, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_24, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_25, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_25, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_26, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_26, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_27, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_27, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_28, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_28, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_29, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_29, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_30, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_30, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_31, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_31, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandPlanted, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandOutput_TOTAL, 0)%></i></strong></td>
			<td align="right"><strong><i><%=FormatNumber(GrandRotasi_Total, 1)%></i></strong></td>
		</tr>
<%End If%>
	</tbody>
</table>

<!--#include file="../Bottom_Report.asp"-->
