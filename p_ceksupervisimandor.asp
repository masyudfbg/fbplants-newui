<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim Name_Report : Name_Report = "CROSSCHECK SUPERVISI VS GAD MANDOR"

Dim v_txtYear, v_txtMonth, v_txtTanggal, v_txtGangCode, v_txtStatus, v_txtTahunTanam, v_txtBlok
Dim ScriptName
Dim rsSupervisiMandor, strSupervisiMandor, strMain, strSupervisiMandor2
Dim NO, cClass, nday, xNum, nSpanTotal

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

v_txtYear		=	Request("txtYear")
v_txtMonth		=	Request("txtMonth")
v_txtTanggal	=	Request("txtTanggal")
v_txtGangCode	=	Request("txtGangCode")
v_txtStatus		=	Request("txtStatus")

If Len(Trim(v_txtYear)) = 0 Then
	v_txtYear = Year(Session("v_Costbook"))
End If

If Len(Trim(v_txtMonth)) = 0 Then
	v_txtMonth = Month(Session("v_Costbook"))
End If

If Len(Trim(v_txtTanggal)) = 0 Then
	v_txtTanggal = Lpad(v_txtMonth, "0", "2") & "/" & GetLastDay(v_txtMonth, v_txtYear) & "/" & v_txtYear
End If

strSupervisiMandor = " select fcba" & _
                     "        , tanggal" & _
                     "        , employee_hist_api.get_gangcode(to_number(to_char(tanggal, 'yyyy')), to_number(to_char(tanggal, 'mm')), fcba, employeecode) gangcode" & _
                     "        , employeecode" & _
                     "        , employeename" & _
                     "        , gad_mandor" & _
                     "        , supervisi" & _
                     "        , decode(gad_mandor + supervisi, 2, 'OK', 'CEK') status" & _
                     " from (" & _
                     " SELECT FCBA" & _
                     "        , TANGGAL" & _
                     "        , ES_CODE EMPLOYEECODE" & _
                     "        , EMPLOYEE_API.GET_EMPLOYEENAME(FCBA, ES_CODE) EMPLOYEENAME" & _
                     "        , SUM(GAD_MANDOR) GAD_MANDOR" & _
                     "        , SUM(SUPERVISI) SUPERVISI" & _
                     " FROM ( SELECT DISTINCT FCBA" & _
                     "               , FDDATE TANGGAL" & _
                     "               , 0 GAD_MANDOR" & _
                     "               , 1 SUPERVISI" & _
                     "               , ES_CODE ES_CODE" & _
                     "        FROM( SELECT FCBA, FDDATE, SUPERVISION_1 ES_CODE" & _
                     "                      FROM ATTENDANCE_GAD" & _
                     "                      WHERE FCBA = '"& Session("FCBA") & "'" & _
					 "                             AND TO_CHAR(fddate, 'yyyymm') = '" & v_txtYear & "' || LPAD(" & v_txtMonth & ", 2, 0)" & _
					 "                             AND attendance IN ('KJ', 'WH', 'WS', 'KK')" & _
					 "                             AND LENGTH(TRIM(SUPERVISION_1)) <> 0" & _
                     "                      UNION ALL" & _
                     "                      SELECT FCBA, FDDATE, SUPERVISION_2 ES_CODE" & _
                     "                      FROM ATTENDANCE_GAD" & _
                     "                      WHERE FCBA = '"& Session("FCBA") & "'" & _
					 "                             AND TO_CHAR(fddate, 'yyyymm') = '" & v_txtYear & "' || LPAD(" & v_txtMonth & ", 2, 0)" & _
					 "                             AND attendance IN ('KJ', 'WH', 'WS', 'KK')" & _
					 "                             AND LENGTH(TRIM(SUPERVISION_2)) <> 0" & _
                     "                      UNION ALL" & _
                     "                      SELECT FCBA, FDDATE, SUPERVISION_3 ES_CODE" & _
                     "                      FROM ATTENDANCE_GAD" & _
                     "                      WHERE FCBA = '"& Session("FCBA") & "'" & _
					 "                             AND TO_CHAR(fddate, 'yyyymm') = '" & v_txtYear & "' || LPAD(" & v_txtMonth & ", 2, 0)" & _
					 "                             AND attendance IN ('KJ', 'WH', 'WS', 'KK')" & _
					 "                             AND LENGTH(TRIM(SUPERVISION_3)) <> 0" & _
                     "                      UNION ALL" & _
                     "                      SELECT FCBA, FDDATE, SUPERVISION_4 ES_CODE" & _
                     "                      FROM ATTENDANCE_GAD" & _
                     "                      WHERE FCBA = '"& Session("FCBA") & "'" & _
					 "                             AND TO_CHAR(fddate, 'yyyymm') = '" & v_txtYear & "' || LPAD(" & v_txtMonth & ", 2, 0)" & _
					 "                             AND attendance IN ('KJ', 'WH', 'WS', 'KK')" & _
					 "                             AND LENGTH(TRIM(SUPERVISION_4)) <> 0" & _
                     "        ) WHERE FCBA = '"& Session("FCBA") & "'" & _
                     "                 AND FDDATE <= TO_DATE('" & v_txtTanggal & "', 'MM/DD/YYYY')" & _
                     "        UNION ALL" & _
                     "        SELECT DISTINCT FCBA" & _
                     "               , FDDATE TANGGAL" & _
                     "               , 1 GAD_MANDOR" & _
                     "               , 0 SUPERVISI" & _
                     "               , EMPLOYEECODE ES_CODE" & _
                     "        FROM ATTENDANCE_GAD" & _
                     "        WHERE ATTENDANCE_GAD_API.GET_GANGTYPECODE (FCBA, FDDATE, EMPLOYEECODE, JOBCODE, LOCATIONTYPE, LOCATIONCODE) = 'SUPERVISION'" & _
                     "              AND ATTENDANCE IN ('KJ', 'WH', 'WS', 'KK')" & _
                     "              AND FCBA = '"& Session("FCBA") & "'" & _
                     "               AND TO_CHAR(fddate, 'yyyymm') = '" & v_txtYear & "' || LPAD(" & v_txtMonth & ", 2, 0)" & _
                     "              AND FDDATE <= TO_DATE('" & v_txtTanggal & "', 'MM/DD/YYYY')" & _
                     " ) WHERE EMPLOYEE_API.GET_GANGTYPECODE(FCBA, ES_CODE) = 'SUPERVISION'"
If Len(Trim(v_txtGangCode)) > 0 Then
	strSupervisiMandor = strSupervisiMandor & " AND EMPLOYEE_API.GET_GANGCODE(FCBA, ES_CODE) = '" & v_txtGangCode & "'"
End If
	strSupervisiMandor = strSupervisiMandor & " GROUP BY FCBA" & _
                                              "          , TANGGAL" & _
                                              "          , ES_CODE" & _
											  "          , EMPLOYEE_API.GET_EMPLOYEENAME(FCBA, ES_CODE)"
If Len(Trim(v_txtStatus)) > 0 Then
	strSupervisiMandor = strSupervisiMandor & " HAVING SUM(GAD_MANDOR) + SUM(SUPERVISI) = '" & v_txtStatus & "'"
End If
	strSupervisiMandor = strSupervisiMandor & " )"

	strSupervisiMandor = strSupervisiMandor & " ORDER BY FCBA" & _
                                              "          , TANGGAL" & _
                                              "          , EMPLOYEE_HIST_API.GET_GANGCODE(TO_NUMBER(TO_CHAR(TANGGAL, 'YYYY')), TO_NUMBER(TO_CHAR(TANGGAL, 'MM')), FCBA, EMPLOYEECODE)" & _
                                              "          , EMPLOYEECODE"

'response.write strSupervisiMandor
%>
<html>

<head>
<title>ePlaS Reporting - <%=PCase(Name_Report)%></title>
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
							<font face="verdana" size="1">TANGGAL</font>
						</td>
						<td align="center">
							<input class="datepicker" type="text" maxlength="10" name="txtTanggal" size="10" value="<%=v_txtTanggal%>" title="Type Date mm/dd/yyyy" >
						</td>
						<td>
							<font face="verdana" size="1">GANG</font>
						</td>
						<td>
							<select name="txtGangCode" >
								<option value = "" <%If Len(Trim(v_txtGangCode)) = 0 then%>  Selected <%End If%>>--ALL--</option>
								<%
								Dim rsGangCode, strGangCode
								strGangCode	=	" SELECT FCCODE, FCNAME FROM GANG WHERE GANGTYPECODE = 'SUPERVISION'" & _
												" AND FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE"
								
								Set rsGangCode = Server.CreateObject("ADODB.Recordset")
								rsGangCode.Open strGangCode, connstr, adOpenStatic 
								
								if not rsGangCode.eof then
								Do While Not rsGangCode.eof
								if rsGangCode("FCCODE") = v_txtGangCode then%>
								<option value="<%=rsGangCode("FCCODE")%>" Selected><%=rsGangCode("FCCODE")%> : <%=rsGangCode("FCNAME")%></option> 
								<%Else%>
								<option value="<%=rsGangCode("FCCODE")%>"><%=rsGangCode("FCCODE")%> : <%=rsGangCode("FCNAME")%></option>  
								<%End if
									rsGangCode.movenext
								loop
								end if
								rsGangCode.Close
								set rsGangCode = Nothing
								%>
							</select>
						</td>
						<td>
							<font face="verdana" size="1">STATUS</font>
						</td>
						<td>
							<select name="txtStatus" >
								<option value = "" <%If Len(Trim(v_txtStatus)) = 0 then%>  Selected <%End If%>>--ALL--</option>
								<option value = "1" <%If v_txtStatus = "1" then%>  Selected <%End If%>>CEK</option>
								<option value = "2" <%If v_txtStatus = "2" then%>  Selected <%End If%>>OK</option>
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
					<h2><%=UCase(Name_Report)%></h2>
					<br/>
					<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
				</div>
			</th>
		</tr>
	<tr>
		<th style="background-color:#00A300;color:#fff;">TANGGAL</th>
		<th style="background-color:#00A300;color:#fff;">GANG</th>
		<th style="background-color:#00A300;color:#fff;">EMPLOYEE CODE</th>
		<th style="background-color:#00A300;color:#fff;">EMPLOYEE NAME</th>
		<th style="background-color:#00A300;color:#fff;">GAD MANDOR</th>
		<th style="background-color:#00A300;color:#fff;">SUPERVISI</th>
		<th style="background-color:#00A300;color:#fff;">STATUS</th>
	</tr>
    </thead>
	<tbody>
<%
Set rsSupervisiMandor = Server.CreateObject("ADODB.Recordset")
rsSupervisiMandor.Open strSupervisiMandor, connstr, adOpenStatic

if rsSupervisiMandor.BOF and rsSupervisiMandor.EOF then
%>
	<tr>
		<td colspan="7">NO DATA AVAILABLE</td>
	</tr>
<%
else
		while not rsSupervisiMandor.eof
%>
				<tr>
					<td align="Right"><%=FbpNum2(rsSupervisiMandor("TANGGAL"))%></td>
					<td align="Center"><%=rsSupervisiMandor("GANGCODE")%></td>
					<td align="Left"><%=rsSupervisiMandor("EMPLOYEECODE")%></td>
					<td align="Left"><%=rsSupervisiMandor("EMPLOYEENAME")%></td>
					<td align="Center"><%=FormatNumber(rsSupervisiMandor("GAD_MANDOR"), 2)%></td>
					<td align="Center"><%=FormatNumber(rsSupervisiMandor("SUPERVISI"), 2)%></td>
					<td align="Center"><%=rsSupervisiMandor("STATUS")%></td>
				</tr>
<%
			rsSupervisiMandor.MoveNext
			
			Wend	
rsSupervisiMandor.Close
set rsSupervisiMandor = Nothing
%>
<%End if%>
	</tbody>
</table>

   </body>
   </html>