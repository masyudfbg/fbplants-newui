<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtYear, v_txtMonth, v_txtGangCode, v_txtGrade
Dim ScriptName
Dim rsEmployee, strEmployee, rsAtt, strAtt, rsKJ, strKJ, status_payroll  
Dim NO, cClass, totalDay
Dim nKJ, nDL
Dim nP1
Dim nP2
Dim nP3
Dim nS1
Dim nS2
Dim nH1
Dim nH2
Dim nCT
Dim nMK
Dim nML
Dim nWH
Dim nWS
Dim nUN
Dim nKK
Dim nKB
Dim nTOTAL


v_txtYear = Request("txtYear")
v_txtMonth = Request("txtMonth")
v_txtGangCode = Request("txtGangCode")
Dim v_txtdetail : v_txtdetail = Request("txtdetail")
Dim txtcolspan : txtcolspan = 1

ScriptName = Request.ServerVariables("SCRIPT_NAME")

if len(trim(v_txtYear)) = 0 then v_txtYear = Year(Session("v_Employee"))
if len(trim(v_txtMonth)) = 0 then v_txtMonth = Month(Session("v_Employee"))

strKJ = " SELECT NVL(payroll_executed, 'FALSE') detail" & _
        " FROM empalldedstatus" & _
		" WHERE fcba = '" & Session ("FCBA") & "'" & _
        "       AND year = " & v_txtYear & "" & _
        "       AND month = " & v_txtMonth & ""

Set rsKJ = Server.CreateObject("ADODB.Recordset")
rsKJ.Open strKJ, connstr, adOpenStatic

IF NOT rsKJ.EOF THEN
	status_payroll = rsKJ("detail")
ELSE
	status_payroll = "FALSE"
END IF
rsKJ.Close
Set rsKJ = Nothing
'response.write status_payroll
IF status_payroll = "TRUE" THEN
	strEmployee = " SELECT *" & _
				  " from (" & _
				  "     SELECT fcba," & _
				  "            nyear," & _
				  "            nmonth," & _
				  "            employee_hist_api.get_gangcode(nyear, nmonth, fcba, employeecode) gangcode," & _
				  "            employeecode," & _
				  "            employee_hist_api.get_fcname(nyear, nmonth, fcba, employeecode) employeename," & _
				  "            employee_hist_api.get_presentgrade(nyear, nmonth, fcba, employeecode) presentgrade," & _
				  "            SUM(DECODE(attendance,'KJ', mandays, 0)) kj," & _
				  "            SUM(DECODE(attendance,'DL', mandays, 0)) dl," & _
				  "            SUM(DECODE(attendance,'P1', mandays, 0)) p1," & _
				  "            SUM(DECODE(attendance,'P2', mandays, 0)) p2," & _
				  "            SUM(DECODE(attendance,'P3', mandays, 0)) p3," & _
				  "            SUM(DECODE(attendance,'S1', mandays, 0)) s1," & _
				  "            SUM(DECODE(attendance,'S2', mandays, 0)) s2," & _
				  "            SUM(DECODE(attendance,'H1', mandays, 0)) h1," & _
				  "            SUM(DECODE(attendance,'H2', mandays, 0)) h2," & _
				  "            SUM(DECODE(attendance,'CT', mandays, 0)) ct," & _
				  "            SUM(DECODE(attendance,'MK', mandays, 0)) mk," & _
				  "            SUM(DECODE(attendance,'ML', mandays, 0)) ml," & _
				  "            SUM(DECODE(attendance,'WH', mandays, 0)) wh," & _
				  "            SUM(DECODE(attendance,'WS', mandays, 0)) ws," & _
				  "            SUM(DECODE(attendance,'UN', mandays, 0)) un,"
				  
	if v_txtdetail = "KK" then
		strEmployee = strEmployee & "            SUM(DECODE(attendance,'KK', mandays,0)) kk,"
	end if

	strEmployee = strEmployee & "            SUM(DECODE(attendance,'KJ', mandays," & _
								"                                  'DL', mandays," & _
								"                                  'P1', mandays," & _
								"                                  'P2', mandays," & _
								"                                  'P3', mandays," & _
								"                                  'S1', mandays," & _
								"                                  'S2', mandays," & _
								"                                  'H1', mandays," & _
								"                                  'H2', mandays," & _
								"                                  'CT', mandays," & _
								"                                  'MK', mandays," & _
								"                                  'ML', mandays," & _
								"                                  'WH', mandays," & _
								"                                  'WS', mandays,"
								
	if v_txtdetail = "KK" then
		strEmployee = strEmployee & "                                  'KK', mandays,"
	end if

		strEmployee = strEmployee & "                                  'UN', mandays, 0)) total" & _
									"     FROM (" & _
									"         SELECT fcba," & _
									"                to_number(to_char(fddate,'YYYY')) nyear," & _
									"                to_number(to_char(fddate,'MM')) nmonth," & _
									"                employeecode," & _
									"                attendance," & _
									"                nvl(sum(mandays), 0) mandays" & _
									"         FROM attendance_gad" & _
									"         WHERE fcba = '" & Session ("FCBA") & "'" & _
									"               AND TO_CHAR(fddate, 'yyyymm') = " & v_txtYear & "||LPAD(" & v_txtMonth & ", 2, 0)" & _
									"         GROUP BY fcba," & _
									"                  to_number(to_char(fddate,'YYYY'))," & _
									"                  to_number(to_char(fddate,'MM'))," & _
									"                  employeecode," & _
									"                  attendance" & _
									"			 UNION ALL " & _
									"		select h.fcba,to_number(to_char(h.bkmdate,'YYYY')) nyear, to_number(to_char(h.bkmdate,'MM')) nmonth, " & _
									"		d.employeecode,d.attendance, nvl(sum(d.mandays), 0) mandays " & _
									"		from bkmhdr h " & _
									"		inner join bkmupahdtl d on (h.bkmno=d.bkmno and h.fcba=d.fcba) " & _
									"		WHERE h.postingflg='1' and h.fcba = '" & Session ("FCBA") & "' AND TO_CHAR(h.bkmdate, 'yyyymm') = " & v_txtYear & "||LPAD(" & v_txtMonth & ", 2, 0)  " & _
									"		GROUP BY h.fcba, to_number(to_char(h.bkmdate,'YYYY')), to_number(to_char(h.bkmdate,'MM')), d.employeecode, d.attendance "  & _
									"     )" & _
									"     GROUP BY fcba," & _
									"              nyear," & _
									"              nmonth," & _
									"              employee_hist_api.get_gangcode(nyear, nmonth, fcba, employeecode)," & _
									"              employeecode," & _
									"              employee_hist_api.get_fcname(nyear, nmonth, fcba, employeecode)," & _
									"              employee_hist_api.get_presentgrade(nyear, nmonth, fcba, employeecode)" & _
									" )" & _
									" WHERE employee_hist_api.get_presentgrade(nyear, nmonth, fcba, employeecode) = 'KHL'"
	if  LEN(TRIM(v_txtGangCode)) > 0 then
		strEmployee	= strEmployee & " AND employee_hist_api.get_gangcode(nyear, nmonth, fcba, employeecode) = '" & v_txtGangCode & "'"
	end if

	strEmployee	= strEmployee & " ORDER BY employeecode ASC"
ELSE
	strEmployee = " SELECT *" & _
				  " from (" & _
				  "     SELECT fcba," & _
				  "            gangcode," & _
				  "            employeecode," & _
				  "            employeename," & _
				  "            presentgrade," & _
				  "            SUM(DECODE(attendance,'KJ', mandays, 0)) kj," & _
				  "            SUM(DECODE(attendance,'DL', mandays, 0)) dl," & _
				  "            SUM(DECODE(attendance,'P1', mandays, 0)) p1," & _
				  "            SUM(DECODE(attendance,'P2', mandays, 0)) p2," & _
				  "            SUM(DECODE(attendance,'P3', mandays, 0)) p3," & _
				  "            SUM(DECODE(attendance,'S1', mandays, 0)) s1," & _
				  "            SUM(DECODE(attendance,'S2', mandays, 0)) s2," & _
				  "            SUM(DECODE(attendance,'H1', mandays, 0)) h1," & _
				  "            SUM(DECODE(attendance,'H2', mandays, 0)) h2," & _
				  "            SUM(DECODE(attendance,'CT', mandays, 0)) ct," & _
				  "            SUM(DECODE(attendance,'MK', mandays, 0)) mk," & _
				  "            SUM(DECODE(attendance,'ML', mandays, 0)) ml," & _
				  "            SUM(DECODE(attendance,'WH', mandays, 0)) wh," & _
				  "            SUM(DECODE(attendance,'WS', mandays, 0)) ws," & _
				  "            SUM(DECODE(attendance,'UN', mandays, 0)) un,"
				  
	if v_txtdetail = "KK" then
		strEmployee = strEmployee & "            SUM(DECODE(attendance,'KK', mandays,0)) kk,"
	end if

	strEmployee = strEmployee & "            SUM(DECODE(attendance,'KJ', mandays," & _
								"                                  'DL', mandays," & _
								"                                  'P1', mandays," & _
								"                                  'P2', mandays," & _
								"                                  'P3', mandays," & _
								"                                  'S1', mandays," & _
								"                                  'S2', mandays," & _
								"                                  'H1', mandays," & _
								"                                  'H2', mandays," & _
								"                                  'CT', mandays," & _
								"                                  'MK', mandays," & _
								"                                  'ML', mandays," & _
								"                                  'WH', mandays," & _
								"                                  'WS', mandays,"
								
	if v_txtdetail = "KK" then
		strEmployee = strEmployee & "                                  'KK', mandays,"
	end if

		strEmployee = strEmployee & "                                  'UN', mandays, 0)) total" & _
									"     FROM (" & _
									"         SELECT ag.fcba," & _
									"                ag.employeecode," & _
									"                em.fcname employeename," & _
									"                em.presentgrade," & _
									"                em.gangcode," & _
									"                ag.attendance," & _
									"                NVL(SUM(ag.mandays), 0) mandays" & _
									"         FROM attendance_gad ag" & _
									"     INNER JOIN employee em" & _
									"     ON (ag.fcba = em.fcba AND ag.employeecode = em.fccode)" & _
									"     WHERE ag.fcba = '" & Session ("FCBA") & "'" & _
									"           AND TO_CHAR(ag.fddate, 'yyyymm') = " & v_txtYear & "||LPAD(" & v_txtMonth & ", 2, 0)" & _
									"           AND em.presentgrade = 'KHL'"
	if LEN(TRIM(v_txtGangCode)) > 0 then
		strEmployee	= strEmployee & "           AND em.gangcode = '" & v_txtGangCode & "'"
	end if
	strEmployee	= strEmployee &     "     GROUP BY ag.fcba," & _
									"              ag.employeecode," & _
									"              em.fcname," & _
									"              em.presentgrade," & _
									"              em.gangcode," & _
									"              ag.attendance" & _
									"     )" & _
									"     GROUP BY fcba," & _
									"              gangcode," & _
									"              employeecode," & _
									"              employeename," & _
									"              presentgrade" & _
									" )"

	strEmployee	= strEmployee & " ORDER BY employeecode ASC"
END IF

if v_txtdetail = "KK" then
	txtcolspan = 21
else
	txtcolspan = 20
end if
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Rekap Absensi KHL</title>
<link rel="stylesheet" href="/css/jquery-ui.css" />
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
</head>

<body>
	<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
	<div id="header-wrapper">
	<!--#include file="menutop.asp"-->
    <table id="form-table" class="centered-margin">
	<form action="<%=ScriptName%>" method="post" name="frmEntry">
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td>
		<% 
			Dim mon, Yr, MinYr, MaxYr, i, MM
			mon = Month(Session("v_Costbook"))
			If len(v_txtMonth)=0 then v_txtMonth=mon

			Yr = Year(Session("v_Costbook"))
			MinYr = 2015
			MaxYr = Yr  + 10
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
			For i = 1 to 12
		%>
				<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
		<%
			Next
		%>
					</select>
				</td>
		<td><font face="verdana" size="1">GANGCODE</font></td><td><select name="txtGangCode" ><option value=""></option>
		<%
		Dim rsGangCode, strGangCode
		strGangCode	=	" SELECT FCCODE, FCNAME FROM GANG WHERE ACTIVATION = 'Y' AND FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE ASC"
		
		Set rsGangCode = Server.CreateObject("ADODB.Recordset")
		rsGangCode.Open strGangCode, connstr, adOpenStatic 
		
		if not rsGangCode.eof then
		Do While Not rsGangCode.eof
		if rsGangCode("FCCODE") = v_txtGangCode then%>
		        <option value="<%=rsGangCode("FCCODE")%>" Selected><%=rsGangCode("FCCODE")%>-<%=rsGangCode("FCNAME")%></option> 
		     <%Else%>
		        <option value="<%=rsGangCode("FCCODE")%>"><%=rsGangCode("FCCODE")%>-<%=rsGangCode("FCNAME")%></option>  
		    <%End if
			rsGangCode.movenext
		loop
		end if
		rsGangCode.Close
		set rsGangCode = Nothing
		%>
		</select></td>
		<td><font face="verdana" size="1">DETAIL</font></td>
		<td>
			<select name="txtdetail" >
				<option value="" <%if len(trim(v_txtdetail)) = 0 then%> Selected <%End If%>>Non KK</option>
		        <option value="KK" <%if v_txtdetail = "KK" then%> Selected <%End If%>>Dengan KK</option>
			</select>
		</td>
		<td><input type="submit" value="View" /></td>
	</tr>
	</form>
</table>
</div>

	<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="<%=txtcolspan%>" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered" style="width:50%; margin:auto;">
		<h2>LAPORAN REKAP HARI KERJA KHL</h2>
		<br/>
		<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
	</div>
	</div>

       </th>
  </tr>
	<tr>
		<th>No</th>
		<th>No Pekerja</th>
		<th>Nama Pekerja</th>
		<th>Status</th>
		<th>KJ</th>
		<th>DL</th>
		<th>P1</th>
		<th>P2</th>
		<th>P3</th>
		<th>S1</th>
		<th>S2</th>
		<th>H1</th>
		<th>H2</th>
		<th>CT</th>
		<th>MK</th>
		<th>ML</th>
		<th>WH</th>
		<th>WS</th>
		<th>UN</th>
		<%if v_txtdetail = "KK" then%>
			<th>KK</th>
		<%end if%>
		<th>Total</th>
	</tr>
</thead>
<tbody>
<%
'response.write strEmployee
Set rsEmployee = Server.CreateObject("ADODB.Recordset")
rsEmployee.Open strEmployee, connstr, adOpenStatic
if rsEmployee.eof and rsEmployee.eof then
%>
	<tr>
		<td colspan="<%=txtcolspan%>">NO DATA AVAILABLE</td>
	</tr>
<%end if
		NO = 1
		nKJ = 0
		nDL = 0
		nP1 = 0
		nP2 = 0
		nP3 = 0
		nS1 = 0
		nS2 = 0
		nH1 = 0
		nH2 = 0
		nCT = 0
		nMK = 0
		nML = 0
		nWH = 0
		nWS = 0
		nUN = 0
		nKK = 0
		nKB = 0
		nTOTAL = 0

		while not rsEmployee.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center"><%=NO%></td>
				<td align="center"><%=rsEmployee("EMPLOYEECODE")%></td>
				<td align="left"><%=rsEmployee("EMPLOYEENAME")%></td>
				<td align="center"><%=rsEmployee("PRESENTGRADE")%></td>
				<td align="right"><%=FbpNum2(rsEmployee("KJ"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("DL"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("P1"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("P2"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("P3"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("S1"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("S2"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("H1"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("H2"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("CT"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("MK"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("ML"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("WH"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("WS"))%></td>
				<td align="right"><%=FbpNum2(rsEmployee("UN"))%></td>
				<%if v_txtdetail = "KK" then%>
					<td align="right"><%=FbpNum2(rsEmployee("KK"))%></td>
				<%end if%>
				<td align="right"><%=FbpNum2(rsEmployee("TOTAL"))%></td>
<%
			NO = NO + 1
			nKJ = nKJ + Cdbl(rsEmployee("KJ"))
			nDL = nDL + Cdbl(rsEmployee("DL"))
			nP1 = nP1 + Cdbl(rsEmployee("P1"))
			nP2 = nP2 + Cdbl(rsEmployee("P2"))
			nP3 = nP3 + Cdbl(rsEmployee("P3"))
			nS1 = nS1 + Cdbl(rsEmployee("S1"))
			nS2 = nS2 + Cdbl(rsEmployee("S2"))
			nH1 = nH1 + Cdbl(rsEmployee("H1"))
			nH2 = nH2 + Cdbl(rsEmployee("H2"))
			nCT = nCT + Cdbl(rsEmployee("CT"))
			nMK = nMK + Cdbl(rsEmployee("MK"))
			nML = nML + Cdbl(rsEmployee("ML"))
			nWH = nWH + Cdbl(rsEmployee("WH"))
			nWS = nWS + Cdbl(rsEmployee("WS"))
			nUN = nUN + Cdbl(rsEmployee("UN"))
			if v_txtdetail = "KK" then
				nKK = nKK + Cdbl(rsEmployee("KK"))
			end if
			nTOTAL = nTOTAL + Cdbl(rsEmployee("TOTAL"))

			rsEmployee.MoveNext
		Wend
rsEmployee.Close
set rsEmployee = Nothing
%>
			<tr bgcolor="#F4F4A0">
				<td align="center" colspan="4"><strong>Total</strong></td>
				<td align="right"><strong><%=FbpNum2(nKJ)%></strong></td>
				<td align="right"><strong><%=FbpNum2(nDL)%></strong></td>
				<td align="right"><strong><%=FbpNum2(nP1)%></strong></td>
				<td align="right"><strong><%=FbpNum2(nP2)%></strong></td>
				<td align="right"><strong><%=FbpNum2(nP3)%></strong></td>
				<td align="right"><strong><%=FbpNum2(nS1)%></strong></td>
				<td align="right"><strong><%=FbpNum2(nS2)%></strong></td>
				<td align="right"><strong><%=FbpNum2(nH1)%></strong></td>
				<td align="right"><strong><%=FbpNum2(nH2)%></strong></td>
				<td align="right"><strong><%=FbpNum2(nCT)%></strong></td>
				<td align="right"><strong><%=FbpNum2(nMK)%></strong></td>
				<td align="right"><strong><%=FbpNum2(nML)%></strong></td>
				<td align="right"><strong><%=FbpNum2(nWH)%></strong></td>
				<td align="right"><strong><%=FbpNum2(nWS)%></strong></td>
				<td align="right"><strong><%=FbpNum2(nUN)%></strong></td>
				<%if v_txtdetail = "KK" then%>
					<td align="right"><strong><%=FbpNum2(nKK)%></strong></td>
				<%end if%>
				<td align="right"><strong><%=FbpNum2(nTOTAL)%></strong></td>
			</tr>
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
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("FCMANAGER")%></u></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="2"><strong><u><%=Session("KTU")%></u></strong></td>
	   </tr>
	   	  <tr> 
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>MANAGER</i></strong></td>
		<td width="50%" align="center"><font face="Verdana" size="1"><strong><i>KTU</i></strong></td>
	   </tr>
   </table>
   </body>
   </html>
