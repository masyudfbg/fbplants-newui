<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim v_txtYear, v_txtMonth, v_txtGangCode, v_txtGrade, v_txtSection
Dim ScriptName
Dim rsEmployee, strEmployee, rsAtt, strAtt, rsKJ, strKJ, status_payroll
Dim NO, cClass, totalDay
Dim nKJ
Dim nDL
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
Dim nKL
Dim nTOTAL

v_txtYear = Request("txtYear")
v_txtMonth = Request("txtMonth")
v_txtGangCode = Request("txtGangCode")
v_txtGrade = Request("txtGrade")
v_txtSection = Request("txtSection")

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

if (v_txtYear = "" or len(trim(v_txtYear)) = 0) then
   v_txtYear = Year(Session("v_Employee"))
end if
if (v_txtMonth = "" or len(trim(v_txtMonth)) = 0) then
   v_txtMonth = Month(Session("v_Employee"))
end if

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
				  "            SUM(DECODE(attendance,'P1', 1, 0)) p1," & _
				  "            SUM(DECODE(attendance,'P2', 1, 0)) p2," & _
				  "            SUM(DECODE(attendance,'P3', 1, 0)) p3," & _
				  "            SUM(DECODE(attendance,'S1', 1, 0)) s1," & _
				  "            SUM(DECODE(attendance,'S2', 1, 0)) s2," & _
				  "            SUM(DECODE(attendance,'H1', 1, 0)) h1," & _
				  "            SUM(DECODE(attendance,'H2', 1, 0)) h2," & _
				  "            SUM(DECODE(attendance,'CT', 1, 0)) ct," & _
				  "            SUM(DECODE(attendance,'MK', 1, 0)) mk," & _
				  "            SUM(DECODE(attendance,'ML', 1, 0)) ml," & _
				  "            SUM(DECODE(attendance,'WH', mandays, 0)) wh," & _
				  "            SUM(DECODE(attendance,'WS', mandays, 0)) ws," & _
				  "            SUM(DECODE(attendance,'UN', 1, 0)) un," & _
				  "            SUM(DECODE(attendance,'KK', mandays,0)) kk," & _
				  "			   SUM (DECODE (attendance, 'KL', 1, 0)) kl, " & _
				  "            SUM(DECODE(attendance,'KJ', mandays," & _
				  "                                  'DL', mandays," & _
				  "                                  'P1', 1," & _
				  "                                  'P2', 1," & _
				  "                                  'P3', 1," & _
				  "                                  'S1', 1," & _
				  "                                  'S2', 1," & _
				  "                                  'H1', 1," & _
				  "                                  'H2', 1," & _
				  "                                  'CT', 1," & _
				  "                                  'MK', 1," & _
				  "                                  'ML', 1," & _
				  "                                  'WH', mandays," & _
				  "                                  'WS', mandays," & _
				  "                                  'KK', mandays," & _
				  "                                  'KL', 1," & _
				  "                                  'UN', 1, 0)) total" & _
				  "     FROM (" & _
				  "         SELECT fcba," & _
				  "                to_number(to_char(fddate,'YYYY')) nyear," & _
				  "                to_number(to_char(fddate,'MM')) nmonth," & _
				  "                fddate," & _
				  "                employeecode," & _
				  "                attendance," & _
				  "                nvl(sum(mandays), 0) mandays" & _
				  "         FROM attendance_gad" & _
				  "         WHERE fcba = '" & Session ("FCBA") & "'" & _
				  "               AND TO_CHAR(fddate, 'yyyymm') = " & v_txtYear & "||LPAD(" & v_txtMonth & ", 2, 0)" & _
				  "         GROUP BY fcba," & _
				  "                  to_number(to_char(fddate,'YYYY'))," & _
				  "                  to_number(to_char(fddate,'MM'))," & _
				  "                  fddate," & _
				  "                  employeecode," & _
				  "                  attendance" & _
				  "         UNION " & _ 
				  "         SELECT a.fcba," & _
				  "                to_number(to_char(bkmdate,'YYYY')) nyear," & _
				  "                to_number(to_char(bkmdate,'MM')) nmonth," & _
				  "                bkmdate fddate," & _
				  "                employeecode," & _
				  "                attendance," & _
				  "                nvl(sum(mandays), 0) mandays" & _
				  "         FROM bkmupahdtl a inner join bkmhdr b ON (a.bkmno=b.bkmno and a.fcba=b.fcba) " & _
				  "         WHERE b.postingflg='1' and a.fcba = '" & Session ("FCBA") & "'" & _
				  "               AND TO_CHAR(bkmdate, 'yyyymm') = " & v_txtYear & "||LPAD(" & v_txtMonth & ", 2, 0)" & _
				  "         GROUP BY a.fcba," & _
				  "                  to_number(to_char(bkmdate,'YYYY'))," & _
				  "                  to_number(to_char(bkmdate,'MM'))," & _
				  "                  bkmdate," & _
				  "                  employeecode," & _
				  "                  attendance" & _
				  
				  "     )" & _
				  "     GROUP BY fcba," & _
				  "              nyear," & _
				  "              nmonth," & _
				  "              employee_hist_api.get_gangcode(nyear, nmonth, fcba, employeecode)," & _
				  "              employeecode," & _
				  "              employee_hist_api.get_fcname(nyear, nmonth, fcba, employeecode)," & _
				  "              employee_hist_api.get_presentgrade(nyear, nmonth, fcba, employeecode)" & _
				  " )" & _
				  " WHERE employee_hist_api.get_presentgrade(nyear, nmonth, fcba, employeecode) IN ('KHT', 'KBT')"
	if  LEN(TRIM(v_txtGrade)) > 0 then
		strEmployee	= strEmployee & " AND employee_hist_api.get_presentgrade(nyear, nmonth, fcba, employeecode) = '" & v_txtGrade & "'"
	end if
	if  LEN(TRIM(v_txtSection)) > 0 then
		strEmployee	= strEmployee & " AND employee_hist_api.get_sectionname(nyear, nmonth, fcba, employeecode) = '" & v_txtSection & "'"
	end if
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
				  "            SUM(DECODE(attendance,'P1', 1, 0)) p1," & _
				  "            SUM(DECODE(attendance,'P2', 1, 0)) p2," & _
				  "            SUM(DECODE(attendance,'P3', 1, 0)) p3," & _
				  "            SUM(DECODE(attendance,'S1', 1, 0)) s1," & _
				  "            SUM(DECODE(attendance,'S2', 1, 0)) s2," & _
				  "            SUM(DECODE(attendance,'H1', 1, 0)) h1," & _
				  "            SUM(DECODE(attendance,'H2', 1, 0)) h2," & _
				  "            SUM(DECODE(attendance,'CT', 1, 0)) ct," & _
				  "            SUM(DECODE(attendance,'MK', 1, 0)) mk," & _
				  "            SUM(DECODE(attendance,'ML', 1, 0)) ml," & _
				  "            SUM(DECODE(attendance,'WH', mandays, 0)) wh," & _
				  "            SUM(DECODE(attendance,'WS', mandays, 0)) ws," & _
				  "            SUM(DECODE(attendance,'UN', 1, 0)) un," & _
				  "            SUM(DECODE(attendance,'KK', mandays,0)) kk," & _
				  "			   SUM (DECODE (attendance, 'KL', 1, 0)) kl, " & _
				  "            SUM(DECODE(attendance,'KJ', mandays," & _
				  "                                  'DL', mandays," & _
				  "                                  'P1', 1," & _
				  "                                  'P2', 1," & _
				  "                                  'P3', 1," & _
				  "                                  'S1', 1," & _
				  "                                  'S2', 1," & _
				  "                                  'H1', 1," & _
				  "                                  'H2', 1," & _
				  "                                  'CT', 1," & _
				  "                                  'MK', 1," & _
				  "                                  'ML', 1," & _
				  "                                  'WH', mandays," & _
				  "                                  'WS', mandays," & _
				  "                                  'KK', mandays," & _
				  "                                  'KL', 1," & _
				  "                                  'UN', 1, 0)) total" & _
				  "     FROM (" & _
				  "         SELECT ag.fcba," & _
				  "                ag.fddate," & _
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
				  "           AND em.presentgrade IN ('KHT', 'KBT')"
	if LEN(TRIM(v_txtGrade)) > 0 then
		strEmployee	= strEmployee & "           AND em.presentgrade = '" & v_txtGrade & "'"
	end if
	if LEN(TRIM(v_txtSection)) > 0 then
		strEmployee	= strEmployee & "           AND em.sectionname = '" & v_txtSection & "'"
	end if
	if LEN(TRIM(v_txtGangCode)) > 0 then
		strEmployee	= strEmployee & "           AND em.gangcode = '" & v_txtGangCode & "'"
	end if
	strEmployee	= strEmployee &     "     GROUP BY ag.fcba," & _
									"              ag.fddate," & _
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

' response.write strEmployee
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Rekap Absensi KHT KBT</title>
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
			For i = 1 to 12
		%>
				<Option Value="<%=i%>" <%If cint(i) = cint(v_txtMonth) then%> Selected <%End If%>><%=Bulan(i)%></option>
		<%
			Next
		%>
					</select>
				</td>
		<td><font face="verdana" size="1">SECTION</font></td>
		<td>
			<select name="txtSection" >
				<option value="">--ALL--</option>
				<%
				Dim rsSection, strSection
				strSection	=	"select p.fccode, p.fcname from parameter p where p.parcode = 'E0002' and fcba = '" & session("fcba") & "' order by p.fcname asc"
				
				Set rsSection = Server.CreateObject("ADODB.Recordset")
				rsSection.Open strSection, connstr, adOpenStatic 
				
				if not rsSection.eof then
				Do While Not rsSection.eof
				if rsSection("FCCODE") = v_txtSection then%>
						<option value="<%=rsSection("FCCODE")%>" Selected><%=rsSection("FCCODE")%> : <%=rsSection("FCNAME")%></option> 
				<%Else%>
						<option value="<%=rsSection("FCCODE")%>"><%=rsSection("FCCODE")%> : <%=rsSection("FCNAME")%></option>  
				<%End if
					rsSection.movenext
				loop
				end if
				rsSection.Close
				set rsSection = Nothing
				%>
				</select>
		</td>
		<td><font face="verdana" size="1">GANGCODE</font></td>
		<td>
			<select name="txtGangCode" >
				<option value="">--ALL--</option>
				<%
				Dim rsGangCode, strGangCode
				strGangCode	=	" SELECT FCCODE, FCNAME FROM GANG WHERE ACTIVATION = 'Y' AND FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE ASC"
				
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
		<td><font face="verdana" size="1">STATUS</font></td>
		<td>
			<select name="txtGrade" >
				<option value="" <%if len(trim(v_txtGrade)) = 0 then%> Selected <%End If%>>--ALL--</option>
		        <option value="KHT" <%if v_txtGrade = "KHT" then%> Selected <%End If%>>KHT : Karyawan Harian Tetap</option>
				<option value="KBT" <%if v_txtGrade = "KBT" then%> Selected <%End If%>>KBT : Karyawan Bulanan Tetap</option> 
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
    <th colspan="23" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered" style="width:50%; margin:auto;">
		<h2>LAPORAN HARI KERJA KHT dan KBT</h2>
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
		<th>Gang</th>
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
		<th>KK</th>
		<th>KL</th>
		<th>Total</th>
	</tr>
</thead>
    <tbody>

<%
Set rsEmployee = Server.CreateObject("ADODB.Recordset")
rsEmployee.Open strEmployee, connstr, adOpenStatic
if rsEmployee.bof and rsEmployee.eof then
%>
	<tr>
		<td colspan="21">NO DATA AVAILABLE</td>
	</tr>
<%end if
		NO = 1
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
				<td align="center"><%=rsEmployee("gangcode")%></td>
				<td align="right"><%=rsEmployee("KJ")%></td>
				<td align="right"><%=rsEmployee("DL")%></td>
				<td align="right"><%=rsEmployee("P1")%></td>
				<td align="right"><%=rsEmployee("P2")%></td>
				<td align="right"><%=rsEmployee("P3")%></td>
				<td align="right"><%=rsEmployee("S1")%></td>
				<td align="right"><%=rsEmployee("S2")%></td>
				<td align="right"><%=rsEmployee("H1")%></td>
				<td align="right"><%=rsEmployee("H2")%></td>
				<td align="right"><%=rsEmployee("CT")%></td>
				<td align="right"><%=rsEmployee("MK")%></td>
				<td align="right"><%=rsEmployee("ML")%></td>
				<td align="right"><%=rsEmployee("WH")%></td>
				<td align="right"><%=rsEmployee("WS")%></td>
				<td align="right"><%=rsEmployee("UN")%></td>
				<td align="right"><%=rsEmployee("KK")%></td>
				<td align="right"><%=rsEmployee("KL")%></td>
				<td align="right"><%=rsEmployee("TOTAL")%></td>
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
			nKK = nKK + Cdbl(rsEmployee("KK"))
			nKL = nKL + Cdbl(rsEmployee("KL"))
			nTOTAL = nTOTAL + Cdbl(rsEmployee("TOTAL"))
			rsEmployee.MoveNext
		Wend
rsEmployee.Close
set rsEmployee = Nothing
%>
			<tr bgcolor="#F4F4A0">
				<td align="center" colspan="5"><strong>TOTAL</strong></td>
				<td align="right"><strong><%=nKJ%></strong></td>
				<td align="right"><strong><%=nDL%></strong></td>
				<td align="right"><strong><%=nP1%></strong></td>
				<td align="right"><strong><%=nP2%></strong></td>
				<td align="right"><strong><%=nP3%></strong></td>
				<td align="right"><strong><%=nS1%></strong></td>
				<td align="right"><strong><%=nS2%></strong></td>
				<td align="right"><strong><%=nH1%></strong></td>
				<td align="right"><strong><%=nH2%></strong></td>
				<td align="right"><strong><%=nCT%></strong></td>
				<td align="right"><strong><%=nMK%></strong></td>
				<td align="right"><strong><%=nML%></strong></td>
				<td align="right"><strong><%=nWH%></strong></td>
				<td align="right"><strong><%=nWS%></strong></td>
				<td align="right"><strong><%=nUN%></strong></td>
				<td align="right"><strong><%=nKK%></strong></td>
				<td align="right"><strong><%=nKL%></strong></td>
				<td align="right"><strong><%=nTOTAL%></strong></td>
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
