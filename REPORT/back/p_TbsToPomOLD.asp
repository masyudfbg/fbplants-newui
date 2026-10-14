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
Dim rsTBS, strTBS, rsAtt, strAtt, rsKJ, strKJ
Dim NO, cClass, totalDay
Dim nAFD1
Dim nAFD2
Dim nAFD3
Dim nAFD4
Dim nAFD5
Dim nAFD6
Dim nAFD7
Dim nAFD8
Dim nH1
Dim nH2
Dim nCT
Dim nMK
Dim nML
Dim nWH
Dim nWS
Dim nUN
Dim nKK
Dim nTOTAL

v_txtYear		=	Request("txtYear")
v_txtMonth		=	Request("txtMonth")

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")

if request("txtYear")="" then
   v_txtYear	=	Year(Session("v_Employee"))
end if
if request("txtMonth")="" then
   v_txtMonth	=	Month(Session("v_Employee"))
end if

strTBS	=	"SELECT FCBA, " & _
            "     RECEPTIONDATE, " & _
            "     AFD_1, " & _
            "     AFD_2, " & _
            "     AFD_3, " & _
            "     AFD_4, " & _
            "     AFD_5, " & _
            "     AFD_6, " & _
            "     AFD_7, " & _
            "     AFD_8, " & _
            "     TOTAL " & _
            "     FROM ( SELECT FCBA, " & _
            "     RECEPTIONDATE, " & _
            "     AFD_1, " & _
            "     AFD_2, " & _
            "     AFD_3, " & _
            "     AFD_4, " & _
            "     AFD_5, " & _
            "     AFD_6, " & _
            "     AFD_7, " & _
            "     AFD_8, " & _
            "     SUM (AFD_1 + AFD_2 + AFD_3 + AFD_4 + AFD_5 + AFD_6 + AFD_7 + AFD_8) TOTAL " & _
            "     FROM ( SELECT harv.fcba, " & _
            "     receptiondate, " & _
            "     harv.productcode, " & _
            "     harv.cropcode, " & _
            "     SUM (DECODE (f.division, 'AFD-1', qty, 0)) afd_1, " & _
            "     SUM (DECODE (f.division, 'AFD-2', qty, 0)) afd_2, " & _
            "     SUM (DECODE (f.division, 'AFD-3', qty, 0)) afd_3, " & _
            "     SUM (DECODE (f.division, 'AFD-4', qty, 0)) afd_4, " & _
            "     SUM (DECODE (f.division, 'AFD-5', qty, 0)) afd_5, " & _
            "     SUM (DECODE (f.division, 'AFD-6', qty, 0)) afd_6, " & _
            "     SUM (DECODE (f.division, 'AFD-7', qty, 0)) afd_7, " & _
            "     SUM (DECODE (f.division, 'AFD-8', qty, 0)) afd_8 " & _
            "     FROM ( SELECT FCBA, " & _
            "     RECEPTIONDATE, " & _
            "     PRODUCTCODE, " & _
            "     CROPCODE, " & _
            "     FIELDCODE, " & _
            "     SUM (MILL_WEIGHT) QTY " & _
            "     FROM harvesting " & _
            "     GROUP BY FCBA, " & _
            "     RECEPTIONDATE, " & _
            "     PRODUCTCODE, " & _
            "     CROPCODE, " & _
            "     FIELDCODE " & _
            "     ORDER BY FCBA, RECEPTIONDATE) harv " & _
            "     INNER JOIN " & _
            "     field f " & _
            "     ON (harv.fieldcode = f.fccode AND harv.fcba = f.fcba) " & _
            "     GROUP BY harv.fcba, " & _
            "     receptiondate, " & _
            "     harv.productcode, " & _
            "     harv.cropcode) " & _
            "     GROUP BY FCBA, " & _
            "     RECEPTIONDATE, " & _
            "     AFD_1, " & _
            "     AFD_2, " & _
            "     AFD_3, " & _
            "     AFD_4, " & _
            "     AFD_5, " & _
            "     AFD_6, " & _
            "     AFD_7, " & _
            "     AFD_8) " & _
			"	  WHERE FCBA = '" & Session ("FCBA") & "' AND TO_NUMBER(TO_CHAR(RECEPTIONDATE, 'MM')) = " & v_txtMonth & _
			" AND TO_NUMBER(TO_CHAR(RECEPTIONDATE, 'YYYY')) = " & v_txtYear & " order by RECEPTIONDATE"
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - TBS ke PKS</title>
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
body{font-family:'Arial Narrow';font-size:12px;}
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
	<div id="header">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered">
		<h2>LAPORAN PENGIRIMAN TBS</h3>
		<br/>
		<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
	</div>
	</div>
<div class="printhide">
<table id="form-table" class="centered-margin">
	<form action="<%=ScriptName%>" method="post" name="frmEntry">
	<tr>
		<td><font face="verdana" size="1">YEAR</font></td>
<% 
	Dim mon, Yr, MinYr, MaxYr, i, MM
    mon = Month(Session("v_Harvesting"))
    If len(v_txtMonth)=0 then v_txtMonth=mon

    Yr = Year(Session("v_Harvesting"))
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
		<td><input type="submit" value="View" /></td>
	</tr>
	</form>
</table>
</div>
<table border="0" width="100%" id="product-table">
	<tr>
		<th style="color:#FFFFFF" rowspan="2" width="10%">TANGGAL</th>
		<th style="color:#FFFFFF" colspan="8" width="10%">PRODUKSI AFDELING</th>
		<th style="color:#FFFFFF" rowspan="2" width="10%">TOTAL</th>
	</tr>
	<tr>
		<th style="color:#FFFFFF" width="10%">I</th>
		<th style="color:#FFFFFF" width="10%">II</th>
		<th style="color:#FFFFFF" width="10%">III</th>
		<th style="color:#FFFFFF" width="10%">IV</th>
		<th style="color:#FFFFFF" width="10%">V</th>
		<th style="color:#FFFFFF" width="10%">VI</th>
		<th style="color:#FFFFFF" width="10%">VII</th>
		<th style="color:#FFFFFF" width="10%">VIII</th>
	</tr>

<%
Set rsTBS = Server.CreateObject("ADODB.Recordset")
rsTBS.Open strTBS, connstr, adOpenStatic
if rsTBS.bof and rsTBS.eof then
%>
	<tr>
		<td colspan="11">NO DATA AVAILABLE</td>
	</tr>
<%end if
		NO = 1
		while not rsTBS.eof
			if no  mod 2 = 1 then
				cClass = "odd"
			else
				cClass = "default"
			end if
%>
			<tr class="<%=cClass%>">
				<td align="center"><%=rsTBS("RECEPTIONDATE")%></td>
				<td align="right"><%=FormatNumber(rsTBS("AFD_1"),0)%></td>
				<td align="right"><%=FormatNumber(rsTBS("AFD_2"),0)%></td>
				<td align="right"><%=FormatNumber(rsTBS("AFD_3"),0)%></td>
				<td align="right"><%=FormatNumber(rsTBS("AFD_4"),0)%></td>
				<td align="right"><%=FormatNumber(rsTBS("AFD_5"),0)%></td>
				<td align="right"><%=FormatNumber(rsTBS("AFD_6"),0)%></td>
				<td align="right"><%=FormatNumber(rsTBS("AFD_7"),0)%></td>
				<td align="right"><%=FormatNumber(rsTBS("AFD_8"),0)%></td>
				<td align="right"><%=FormatNumber(rsTBS("TOTAL"),0)%></td>
<%
			NO = NO + 1
			nAFD1 = nAFD1 + Cdbl(rsTBS("AFD_1"))
			nAFD2 = nAFD2 + Cdbl(rsTBS("AFD_2"))
			nAFD3 = nAFD3 + Cdbl(rsTBS("AFD_3"))
			nAFD4 = nAFD4 + Cdbl(rsTBS("AFD_4"))
			nAFD5 = nAFD5 + Cdbl(rsTBS("AFD_5"))
			nAFD6 = nAFD6 + Cdbl(rsTBS("AFD_6"))
			nAFD7 = nAFD7 + Cdbl(rsTBS("AFD_7"))
			nAFD8 = nAFD8 + Cdbl(rsTBS("AFD_8"))
			nTOTAL = nTOTAL + Cdbl(rsTBS("TOTAL"))
			rsTBS.MoveNext
		Wend
rsTBS.Close
set rsTBS = Nothing
%>
			<tr bgcolor="#F4F4A0">
				<td align="center" colspan="1"><strong>TOTAL</td>
				<td align="right"><strong><%=FormatNumber(nAFD1,0)%></strong></td>
				<td align="right"><strong><%=FormatNumber(nAFD2,0)%></strong></td>
				<td align="right"><strong><%=FormatNumber(nAFD3,0)%></strong></td>
				<td align="right"><strong><%=FormatNumber(nAFD4,0)%></strong></td>
				<td align="right"><strong><%=FormatNumber(nAFD5,0)%></strong></td>
				<td align="right"><strong><%=FormatNumber(nAFD6,0)%></strong></td>
				<td align="right"><strong><%=FormatNumber(nAFD7,0)%></strong></td>
				<td align="right"><strong><%=FormatNumber(nAFD8,0)%></strong></td>
				<td align="right"><strong><%=FormatNumber(nTOTAL,0)%></strong></td>
			</tr>
</table>
</div>
</body>
</html>
