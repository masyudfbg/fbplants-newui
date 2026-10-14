<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<%
Dim rsJob, strJob, rsSourceType, strSourceType, rsCA, strCA
Dim ScriptName, v_txtYear, v_txtMonth, ZeroValue, v_txtJobCode
Dim NO, cClass, nSubTotal, nGrandTotal
v_txtYear	=	Request("txtYear")
v_txtMonth	=	Request("txtMonth")
v_txtJobCode=	Request("txtJobCode")

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
ZeroValue	=	0
nGrandTotal	=	0

    if request("txtYear")="" then
	   v_txtYear=Year(Now)
	end if
    if request("txtMonth")="" then
	   v_txtMonth=Month(Now)
	end if

strJob		= 	" SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '" & Session("FCBA") & "' AND FCCODE = '" & v_txtJobCode & "'"
strJob		=	strJob & " ORDER BY FCCODE ASC"

Set rsJob = Server.CreateObject("ADODB.Recordset")
rsJob.Open strJob, connstr, adOpenStatic 

%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPS Reporting - Nota Debet</title>
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
	<div id="header-wrapper">
	<!--#include file="menutop.asp"-->
	<div id="header">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered">
		<h2>NOTA DEBET - CREDIT</h3>
		<h5>PERIODE&nbsp;:&nbsp;<%=textMonth(v_txtMonth)%>&nbsp;<%=v_txtYear%></h5>
	</div>
	</div>
<table id="form-table" class="centered-margin">
	<form action="<%=ScriptName%>" method="post" name="frmEntry">
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td><td width="4"><input type="text" name="txtYear" width="4" value="<%=v_txtYear%>"></td>
		<td><font face="verdana" size="1">MONTH</font></td><td width="2"><input type="text" name="txtMonth" width="2" value="<%=v_txtMonth%>"></td>
		<td><select name="txtJobCode" >
		<%
		Dim rsJobCat, strJobCat
		strJobCat	=	" SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '" & Session("FCBA") & "'" & _
						" AND UPPER(FCNAME) LIKE 'INTRACO%'"
		strJobCat	=	strJobCat & " ORDER BY FCCODE ASC"
		
		Set rsJobCat = Server.CreateObject("ADODB.Recordset")
		rsJobCat.Open strJobCat, connstr, adOpenStatic 
		
		if not rsJobCat.eof then
		Do While Not rsJobCat.eof
		if rsJobCat("FCCODE") = v_txtJobCode then%>
		        <option value="<%=rsJobCat("FCCODE")%>" Selected><%=rsJobCat("FCCODE")%>-<%=rsJobCat("FCNAME")%></option> 
		     <%Else%>
		        <option value="<%=rsJobCat("FCCODE")%>"><%=rsJobCat("FCCODE")%>-<%=rsJobCat("FCNAME")%></option>  
		    <%End if
			rsJobCat.movenext
		loop
		end if
		rsJobCat.Close
		set rsJobCat = Nothing
		%>
		</select></td>
		<td><input type="submit" value="View" /></td>
	</tr>
	</form>
</table>
<%if not rsJob.eof then
	while not rsJob.eof%>
<table border="0" width="100%" id="#"><tr><td align="left"><%=rsJob("FCCODE")%>&nbsp;<%=rsJob("FCNAME")%></td></tr></table>
<%
		rsJob.MoveNext
	Wend
	rsJob.Close
	set rsJob = Nothing
end if
%>
<table border="0" width="100%" id="product-table">
	<tr>
		<th rowspan="2">NO</th>
		<th rowspan="2">TANGGAL</th>
		<th rowspan="2">NO. BUKTI</th>
		<th rowspan="2">REFERENSI</th>
		<th rowspan="2">KETERANGAN</th>
		<th colspan="2">JUMLAH (Rp.)</th>
	</tr>
	<tr>
		<th>DEBET</th>
		<th>CREDIT</th>
	</tr>
<%
	Dim rsOpeningBalance, strOpeningBalance
	strOpeningBalance	= 	" SELECT NVL(OPENING_BALANCE,0) OPENING_BALANCE FROM TRIALBALANCE" & _
							" WHERE FCBA = '" & Session("FCBA") & "' AND FCCODE = '" & v_txtJobCode & "'" & _
							" AND PERIOD_YEAR = " & Cint(v_txtYear) & _
							" AND PERIOD_MONTH = " & Cint(v_txtMonth)

Set rsOpeningBalance = Server.CreateObject("ADODB.Recordset")
rsOpeningBalance.Open strOpeningBalance, connstr, adOpenStatic 

%>
<%if rsOpeningBalance.bof and rsOpeningBalance.eof then%>
	<tr><td align="left" colspan="7">NO DATA AVAILABLE</td></tr>
<%else
	nGrandTotal	=	nGrandTotal	+	Cdbl(rsOpeningBalance("OPENING_BALANCE"))
	while not rsOpeningBalance.eof%>
	<tr bgcolor="#9CB770">
		<td align="center" colspan="5"><strong><i>Saldo Awal</i></strong></td>
		<td align="center" colspan="2"><strong><i><%=FormatNumber(rsOpeningBalance("OPENING_BALANCE"),2)%></i></strong></td>
	</tr>
<%
rsOpeningBalance.MoveNext
Wend
rsOpeningBalance.Close
set rsOpeningBalance = Nothing
end if
%>
<%
	strSourceType = 	" SELECT 'BA' SOURCETYPE, 'BANK' SOURCENAME, 1 SEQ FROM DUAL" & _
						" UNION" & _
						" SELECT 'PR' SOURCETYPE, 'PURCHASING' SOURCENAME, 2 SEQ FROM DUAL" & _
						" UNION" & _
						" SELECT 'ST' SOURCETYPE, 'MATERIAL' SOURCENAME, 3 SEQ FROM DUAL" & _
						" UNION" & _
						" SELECT 'EM' SOURCETYPE, 'EMPLOYEE' SOURCENAME, 4 SEQ FROM DUAL" & _
						" UNION" & _
						" SELECT 'WS' SOURCETYPE, 'WORKSHOP' SOURCENAME, 5 SEQ FROM DUAL" & _
						" UNION" & _
						" SELECT 'VH' SOURCETYPE, 'VEHICLE' SOURCENAME, 6 SEQ FROM DUAL" & _
						" UNION" & _
						" SELECT 'MN' SOURCETYPE, 'NURSERY' SOURCENAME, 7 SEQ FROM DUAL ORDER BY SEQ ASC"

Set rsSourceType = Server.CreateObject("ADODB.Recordset")
rsSourceType.Open strSourceType, connstr, adOpenStatic 

%>
<%if rsSourceType.bof and rsSourceType.eof then%>
	<tr><td align="left" colspan="7">NO DATA AVAILABLE</td></tr>
<%else
	while not rsSourceType.eof%>
	<tr><td align="left" colspan="7"bgcolor="F0F0B0"><strong><i><%=rsSourceType("SOURCENAME")%><i></strong></td></tr>

<%
	if rsSourceType("SOURCETYPE") = Cstr("BA") then
	strCA	=	" SELECT" & _
				" DATACA.FCBA FCBA," & _
				" DATACA.TDATE TDATE," & _
				" DATACA.JOBCODE JOBCODE," & _
				" DATACA.SOURCETYPE SOURCETYPE," & _
				" BA.FCNAME REFERENCE," & _
				" DATACA.REFERENCENO REFERENCENO," & _
				" PVD.REMARKS REMARKS," & _
				" NVL (DATACA.QUANTITY, 0) QUANTITY," & _
				" round(NVL (DATACA.AMOUNT, 0),2) AMOUNT " & _
				" FROM" & _
				" COST_ACTIVITY DATACA," & _
				" PAYMENTVOUCHER PV," & _
				" PAYMENTVOUCHERDETAIL PVD," & _
				" BANK BA" & _
				" WHERE" & _
				" (DATACA.FCBA = PV.FCBA AND DATACA.REFERENCENO = PV.VOUCHERCODE AND DATACA.TDATE = PV.DATECREATED AND DATACA.FCBA = BA.FCBA AND DATACA.SOURCECODE = BA.FCCODE)" & _
				" AND" & _
				" (DATACA.FCBA = PVD.FCBA AND DATACA.REFERENCENO = PVD.VOUCHERCODE AND DATACA.JOBCODE = PVD.JOBCODE)" & _
				" AND" & _
				" (PV.FCBA = PVD.FCBA AND PV.VOUCHERCODE = PVD.VOUCHERCODE)" & _
				" AND DATACA.SOURCETYPE = 'BA'" & _
				" AND DATACA.FCBA = '" & Session("FCBA") & "' AND DATACA.JOBCODE = '" & v_txtJobCode & "'"
	elseif rsSourceType("SOURCETYPE") = Cstr("PR") then
	strCA	=	" SELECT" & _
				" DATACA.FCBA FCBA," & _
				" DATACA.TDATE TDATE," & _
				" DATACA.JOBCODE JOBCODE," & _
				" DATACA.SOURCETYPE SOURCETYPE," & _
				" DATACA.SOURCECODE_STATUS REFERENCE," & _
				" DATACA.SOURCECODE REFERENCENO," & _
				" DATACA.REFERENCENO REMARKS," & _
				" NVL (DATACA.QUANTITY, 0) QUANTITY," & _
				" NVL (DATACA.AMOUNT, 0) AMOUNT" & _
				" FROM" & _
				" COST_ACTIVITY DATACA" & _
				" WHERE" & _
				" DATACA.SOURCETYPE = 'PR'" & _
				" AND DATACA.FCBA = '" & Session("FCBA") & "' AND DATACA.JOBCODE = '" & v_txtJobCode & "'"
	elseif rsSourceType("SOURCETYPE") = Cstr("ST") then
	strCA	=	" SELECT" & _
				" DATACA.FCBA FCBA," & _
				" DATACA.TDATE TDATE," & _
				" DATACA.JOBCODE JOBCODE," & _
				" DATACA.SOURCETYPE SOURCETYPE," & _
				" STK.FCNAME REFERENCE," & _
				" DATACA.REFERENCENO REFERENCENO," & _
				" SIVD.REMARKS REMARKS," & _
				" NVL(DATACA.QUANTITY, 0) QUANTITY," & _
				" round(DECODE(NVL(DATACA.QUANTITY,0),0,0,(NVL(DATACA.AMOUNT,0)/NVL(DATACA.QUANTITY,0)) * NVL(SIVD.QUANTITY,0)),2) AMOUNT " & _
				" FROM COST_ACTIVITY DATACA," & _
				" ISSUE SIV," & _
				" ISSUEDETAILS SIVD," & _
				" STOCK STK" & _
				" WHERE" & _
				" (DATACA.FCBA = SIV.FCBA AND DATACA.REFERENCENO = SIV.SIVCODE AND DATACA.TDATE = SIV.SIVDATE)" & _
				" AND" & _
				" (DATACA.FCBA = SIVD.FCBA AND DATACA.REFERENCENO = SIVD.SIVCODE AND DATACA.JOBCODE = SIVD.JOBCODE)" & _
				" AND" & _
				" (SIV.FCBA = SIVD.FCBA AND SIV.SIVCODE = SIVD.SIVCODE)" & _
				" AND" & _
				" (SIVD.FCBA = STK.FCBA AND SIVD.STOCKCODE = STK.FCCODE)" & _
				" AND DATACA.SOURCETYPE = 'ST' " & _
				" AND DATACA.FCBA = '" & Session("FCBA") & "' AND DATACA.JOBCODE = '" & v_txtJobCode & "'"
	elseif rsSourceType("SOURCETYPE") = Cstr("WS") then
	strCA	=	" SELECT" & _
				" DATACA.FCBA FCBA," & _
				" DATACA.TDATE TDATE," & _
				" DATACA.JOBCODE JOBCODE," & _
				" DATACA.SOURCETYPE SOURCETYPE," & _
				" WA.TIME REFERENCENO," & _
				" EM.FCNAME REFERENCE," & _
				" WA.REMARKS REMARKS," & _
				" NVL (DATACA.QUANTITY, 0) QUANTITY," & _
				" round(DECODE(NVL(DATACA.QUANTITY, 0),0,0,(NVL (DATACA.AMOUNT, 0) / NVL (DATACA.QUANTITY, 0)) * NVL (WA.UNIT, 0)),2) AMOUNT " & _
				" FROM COST_ACTIVITY DATACA," & _
				" WORKSHOPACTIVITY WA," & _
				" WORKSHOP EM" & _
				" WHERE" & _
				" (" & _
				" DATACA.FCBA = WA.FCBA" & _
				" AND DATACA.SOURCECODE = WA.WORKSHOPCODE" & _
				" AND DATACA.TDATE = WA.WORKSHOPDATE" & _
				" AND DATACA.JOBCODE = WA.JOB" & _
				" )" & _
				" AND" & _
				" (WA.FCBA = EM.FCBA AND WA.WORKSHOPCODE = EM.FCCODE)" & _
				" AND DATACA.SOURCETYPE = 'WS'" & _
				" AND DATACA.FCBA = '" & Session("FCBA") & "' AND DATACA.JOBCODE = '" & v_txtJobCode & "'"
	elseif rsSourceType("SOURCETYPE") = Cstr("VH") then
	strCA	=	" SELECT" & _
				" DATACA.FCBA FCBA," & _
				" DATACA.TDATE TDATE," & _
				" DATACA.JOBCODE JOBCODE," & _
				" DATACA.SOURCETYPE SOURCETYPE," & _
				" VA.TIME REFERENCENO," & _
				" VH.FCNAME REFERENCE," & _
				" VA.REMARKS REMARKS," & _
				" NVL (DATACA.QUANTITY, 0) QUANTITY," & _
				" ROUND(DECODE(NVL (DATACA.QUANTITY, 0),0,0,(NVL (DATACA.AMOUNT, 0) / NVL (DATACA.QUANTITY, 0)) * NVL (VA.UNIT, 0)),2) AMOUNT" & _
				" FROM COST_ACTIVITY DATACA," & _
				" VEHICLEACTIVITY VA," & _
				" VEHICLE VH" & _
				" WHERE" & _
				" (" & _
				" DATACA.FCBA = VA.FCBA" & _
				" AND DATACA.SOURCECODE = VA.VEHICLECODE" & _
				" AND DATACA.TDATE = VA.VEHICLEDATE" & _
				" AND DATACA.JOBCODE = VA.JOBCODE" & _
				" )" & _
				" AND" & _
				" (VA.FCBA = VH.FCBA AND VA.VEHICLECODE = VH.FCCODE)" & _
				" AND DATACA.SOURCETYPE = 'VH'" & _
				" AND DATACA.FCBA = '" & Session("FCBA") & "' AND DATACA.JOBCODE = '" & v_txtJobCode & "'"
	elseif rsSourceType("SOURCETYPE") = Cstr("MN") then
	strCA	=	" SELECT" & _
				" DATACA.FCBA FCBA," & _
				" DATACA.TDATE TDATE," & _
				" DATACA.JOBCODE JOBCODE," & _
				" DATACA.SOURCETYPE SOURCETYPE," & _
				" DATACA.SOURCECODE REFERENCENO," & _
				" NR.DESCRIPTION REFERENCE," & _
				" NI.REMARKS REMARKS," & _
				" NVL (DATACA.QUANTITY, 0) QUANTITY," & _
				" ROUND(DECODE(NVL (DATACA.QUANTITY, 0),0, (NVL (DATACA.AMOUNT, 0) / NVL (DATACA.QUANTITY, 0)) * NVL(NI.QUANTITY,0)),2) AMOUNT" & _
				" FROM COST_ACTIVITY DATACA," & _
				" NURSERYISSUE NI," & _
				" NURSERY NR" & _
				" WHERE" & _
				" (" & _
				" DATACA.FCBA = NI.FCBA" & _
				" AND DATACA.SOURCECODE = NI.NURSERYCODE" & _
				" AND DATACA.TDATE = NI.ISSUEDATE" & _
				" AND DATACA.JOBCODE = NI.JOBCODE" & _
				" AND DATACA.REFERENCENO = NI.REFERENCE" & _
				" )" & _
				" AND" & _
				" (NI.FCBA = NR.FCBA AND NI.NURSERYCODE = NR.NURSERYCODE)" & _
				" AND DATACA.SOURCETYPE = 'MN'" & _
				" AND DATACA.FCBA = '" & Session("FCBA") & "' AND DATACA.JOBCODE = '" & v_txtJobCode & "'"
	elseif rsSourceType("SOURCETYPE") = Cstr("EM") then
	strCA	=	" SELECT FCBA, TDATE, JOBCODE, SOURCETYPE, REFERENCENO, REFERENCE, REMARKS, AMOUNT FROM (" & _
				" SELECT" & _
				" FCBA, TDATE, JOBCODE, SOURCETYPE, REFERENCENO, REFERENCE, REMARKS, ROUND(SUM(AMOUNT),2) AMOUNT" & _
				" FROM (" & _
				" SELECT" & _
				" EC.FCBA FCBA," & _
				" TRUNC(TRANSDATE,'MM') TDATE," & _
				" JOBCODE," & _
				" SOURCETYPE," & _
				" EM.FCNAME REFERENCENO," & _
				" AD.FCNAME REFERENCE," & _
				" AG.REMARKS REMARKS," & _
				" NVL(AMOUNT,0) AMOUNT" & _
				" FROM EMPLOYEECB EC" & _
				" INNER JOIN ALLOWDED AD ON (EC.FCBA = AD.FCBA AND substr(EC.REFERENCENO,1,3) = AD.FCCODE)" & _
				" INNER JOIN EMPLOYEE EM ON (EC.FCBA = EM.FCBA AND EC.EMPCODE = EM.FCCODE)" & _
				" INNER JOIN ATTENDANCE_GAD AG ON (EC.FCBA = AG.FCBA AND EC.JOBCODE = AG.JOBCODE AND EC.EMPCODE = AG.EMPLOYEECODE AND EC.TRANSDATE = AG.FDDATE)" & _
				" WHERE EC.SOURCETYPE IN ('EM','ES')" & _
				" ) GROUP BY FCBA, TDATE, JOBCODE, SOURCETYPE, REFERENCENO, REFERENCE, REMARKS" & _
				" ) DATACA" & _
				" WHERE DATACA.FCBA = '" & Session("FCBA") & "' AND DATACA.JOBCODE = '" & v_txtJobCode & "'"
	else	
	strCA	=	" SELECT * FROM" & _
				" (SELECT" & _
				" 'DUMMY' FCBA," & _
				" TO_DATE('11/1/2012','MM/DD/YYYY') TDATE," & _
				" 'DUMMY' JOBCODE," & _
				" 'DUMMY' SOURCETYPE," & _
				" 'DUMMY' REFERENCENO," & _
				" 'DUMMY' REFERENCENO," & _
				" 'DUMMY' REMARKS," & _
				" 0 QUANTITY," & _
				" round(0,2) AMOUNT" & _
				" FROM DUAL" & _
				" ) DATACA" & _
				" WHERE" & _
				" DATACA.FCBA = '" & Session("FCBA") & "' AND DATACA.JOBCODE = '" & v_txtJobCode & "'"
	end if
	
	if  LEN(TRIM(v_txtYear)) > 0 then
		strCA	=	strCA & " AND TO_NUMBER(TO_CHAR(DATACA.TDATE,'YYYY')) =" & Cint(v_txtYear)
	end if
	if  LEN(TRIM(v_txtMonth)) > 0 then
		strCA	=	strCA &  " AND TO_NUMBER(TO_CHAR(DATACA.TDATE,'MM')) =" & Cint(v_txtMonth)
	end if
	'response.write strCA
	strCA	=	strCA & " ORDER BY DATACA.TDATE, DATACA.REFERENCENO"
Set rsCA = Server.CreateObject("ADODB.Recordset")
rsCA.Open strCA, connstr, adOpenStatic 
%>
<%if rsCA.bof and rsCA.eof then%>
	<tr><td align="left" colspan="7">NO DATA AVAILABLE</td></tr>
<%else
	if Not rsCA.eof then
	NO = 1
	nSubTotal = 0
	while not rsCA.eof
	if NO  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
	%>
	<tr class="<%=cClass%>">
		<td align="left"><%=NO%></td>
		<td align="left"><%=rsCA("TDATE")%></td>
		<td align="left"><%=rsCA("REFERENCENO")%></td>
		<td align="left"><%=rsCA("REFERENCE")%></td>
		<td align="left"><%=rsCA("REMARKS")%></td>		
		<%if trim(rsCA("AMOUNT")) <> "" then %>
			<%if Cdbl(rsCA("AMOUNT")) > 0 then%>
			<td align="right"><%=FormatNumber(rsCA("AMOUNT"),2)%></td>
			<td align="right"><%=FormatNumber(Cint(ZeroValue),2)%></td>
			<%elseif Cdbl(rsCA("AMOUNT")) < 0 then%>
			<td align="right"><%=FormatNumber(Cint(ZeroValue),2)%></td>
			<td align="right"><%=FormatNumber(rsCA("AMOUNT"),2)%></td>
			<%else%>
			<td align="right"><%=FormatNumber(Cint(ZeroValue),2)%></td>
			<td align="right"><%=FormatNumber(Cint(ZeroValue),2)%></td>
			<%end if%>	
		<%else%>
			<td align="right"><%=FormatNumber(Cint(ZeroValue),2)%></td>
			<td align="right"><%=FormatNumber(Cint(ZeroValue),2)%></td>
		<%end if%>			
	</tr>
<%
		NO = NO + 1 

		if rsCA("AMOUNT") <> "" then 
			nSubTotal = nSubTotal + Cdbl(rsCA("AMOUNT"))
		else
			nSubTotal = nSubTotal + 0
		end if			
		
		rsCA.MoveNext
	Wend
	nGrandTotal = nGrandTotal + nSubTotal	
%>
	<tr bgcolor="#9CB770">
		<td align="center" colspan="5"><strong><i>Sub Total</i></strong></td>
		<%if Cdbl(nSubTotal) > 0 then%>
		<td align="right"><%=FormatNumber(nSubTotal,2)%></td>
		<td align="right"><%=FormatNumber(Cint(ZeroValue),2)%></td>
		<%elseif Cdbl(nSubTotal) < 0 then%>
		<td align="right"><%=FormatNumber(Cint(ZeroValue),2)%></td>
		<td align="right"><%=FormatNumber(nSubTotal,2)%></td>
		<%else%>
		<td align="right"><%=FormatNumber(Cint(ZeroValue),2)%></td>
		<td align="right"><%=FormatNumber(Cint(ZeroValue),2)%></td>
		<%end if%>
	</tr>
<%
end if
	rsCA.Close
	set rsCA = Nothing
end if
		rsSourceType.MoveNext
	Wend
	rsSourceType.Close
	set rsSourceType = Nothing
end if
%>
	<tr bgcolor="#9CB770">
		<td align="center" colspan="5"><strong><i>Grand Total</i></strong></td>
		<%if Cdbl(nGrandTotal) > 0 then%>
		<td align="right"><%=FormatNumber(nGrandTotal,2)%></td>
		<td align="right"><%=FormatNumber(Cint(ZeroValue),2)%></td>
		<%elseif Cdbl(nGrandTotal) < 0 then%>
		<td align="right"><%=FormatNumber(Cint(ZeroValue),2)%></td>
		<td align="right"><%=FormatNumber(nGrandTotal,2)%></td>
		<%else%>
		<td align="right"><%=FormatNumber(Cint(ZeroValue),2)%></td>
		<td align="right"><%=FormatNumber(Cint(ZeroValue),2)%></td>
		<%end if%>
	</tr>
</table>
<!--#include file="../Bottom_Report.asp"-->