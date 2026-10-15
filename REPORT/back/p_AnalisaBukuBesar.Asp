<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, v_txtJobCode
Dim ScriptName
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
v_txtJobCode=	trim(request("txtJobCode"))
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - Analisa Buku Besar</title>
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
	<table id="form-table" class="centered-margin">
		<form name="frmBTL" method="post" action="<%=ScriptName%>">
			<tr>
				<td><font face="verdana" size="1">TAHUN</font></td>
				<td>
<% 
    MM=Request.Form("vMonth")
    yyyy=Request.Form("vYear")

    mon = Month(Session("v_Costbook"))
    If len(MM)=0 then MM=mon

    Yr = Year(Session("v_Costbook"))
    MinYr = Yr - 2
    MaxYr = Yr  + 2
    If len(yyyy)= 0 then yyyy = Yr
%>
				<select name="vYear">
<%
    For i = MinYr to MaxYr
%>
        <Option Value="<%=i%>" <%If cint(i) = cint(yyyy) then%> Selected <%End If%>><%=i%></option>
<%
    Next
%>
				</select>
				</td>
				<td><font face="verdana" size="1">BULAN</font></td>
				<td>
					<select name="vMonth">
<%
    For i = 1 to 12
%>
        <Option Value="<%=i%>" <%If cint(i) = cint(MM) then%> Selected <%End If%>><%=Bulan(i)%></option>
<%
    Next
%>
					</select>
				</td>

				<td><font face="verdana" size="1">JOBCODE</font></td><td><select name="txtJobCode" ><option value=""></option>
				<%
				Dim rsJobCode, strJobCode
				strJobCode	=	" SELECT FCCODE, FCNAME FROM JOB WHERE FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE ASC"
				
				Set rsJobCode = Server.CreateObject("ADODB.Recordset")
				rsJobCode.Open strJobCode, connstr, adOpenStatic 
				
				if not rsJobCode.eof then
				Do While Not rsJobCode.eof
				if rsJobCode("FCCODE") = v_txtJobCode then%>
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
				</select></td>
				<td><input type="submit" value="View" /></td>
			</tr>
		</form>
	</table>
</div>
	<div id="header">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered">
		<h3>LAPORAN BUKU BESAR</h3>
		<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(mm)%>&nbsp;<%=yyyy%></h4>
	</div>
	</div>
<table border="0" width="100%" id="product-table">
	<tr>
		<th style="color:#FFFFFF">KODE PEKERJAAN</th>
		<th style="color:#FFFFFF">JENIS PEKERJAAN</th>
		<th style="color:#FFFFFF">SOURCE TYPE</th>
		<th style="color:#FFFFFF">SOURCE CODE</th>
		<th style="color:#FFFFFF">TARGET TYPE</th>
		<th style="color:#FFFFFF">TARGET CODE</th>
		<th style="color:#FFFFFF">AMOUNT</th>
		<th style="color:#FFFFFF">KETERANGAN</th>

	</tr>
<%
Dim rsAnalisaBukuBesar, strAnalisaBukuBesar, rowno
rowno		=	0
strAnalisaBukuBesar		=	" SELECT TO_NUMBER(TO_CHAR (a.tdate, 'YYYY')) theYear," & _
							" TO_NUMBER(TO_CHAR (a.tdate, 'MM')) theMonth," & _
							" B.JOBGROUPCODE," & _
							" a.jobcode," & _
							" b.fcname jobname," & _
							" a.sourcetype," & _
							" a.sourcecode," & _
							" a.targettype," & _
							" a.targetcode," & _
							" a.amount," & _
							" ADDITIONAL_REPORT_API.getRemarks (a.SourceType, a.SourceCode, a.TargetType, a.TargetCode, a.ReferenceNo, a.fcba)" & _
							" Remarks," & _
							" a.fcba" & _
							" FROM cost_activity a, job b" & _
							" WHERE A.JOBCODE = B.FCCODE AND a.fcba = b.fcba" & _
							" AND a.FCBA = '" & Session("FCBA") & "'"

if  LEN(TRIM(yyyy)) > 0 then
	strAnalisaBukuBesar	=	strAnalisaBukuBesar & " AND TO_NUMBER(TO_CHAR (a.tdate, 'YYYY')) = " & yyyy
end if
if  LEN(TRIM(MM)) > 0 then
	strAnalisaBukuBesar	=	strAnalisaBukuBesar & " AND TO_NUMBER(TO_CHAR (a.tdate, 'MM')) = " & MM
end if
if  LEN(TRIM(v_txtJobCode)) > 0 then
	strAnalisaBukuBesar	=	strAnalisaBukuBesar & " AND JOBCODE = '" & v_txtjobCode & "'"
end if

Set rsAnalisaBukuBesar = Server.CreateObject("ADODB.Recordset")
rsAnalisaBukuBesar.Open strAnalisaBukuBesar, connstr, adOpenStatic
if Not rsAnalisaBukuBesar.eof then

	While Not rsAnalisaBukuBesar.eof
			rowno	=	rowno	+	1
%>
	<tr>
		<td align="center"><%=rsAnalisaBukuBesar("JOBCODE")%></td>
		<td align="left"><%=rsAnalisaBukuBesar("JOBNAME")%></td>
		<td align="center"><%=rsAnalisaBukuBesar("SOURCETYPE")%></td>
		<td align="left"><%=rsAnalisaBukuBesar("SOURCECODE")%></td>
		<td align="center"><%=rsAnalisaBukuBesar("TARGETTYPE")%></td>
		<td align="left"><%=rsAnalisaBukuBesar("TARGETCODE")%></td>
		<td align="right"><%=FormatNumber(rsAnalisaBukuBesar("AMOUNT"),2)%></td>
		<td align="left"><%=rsAnalisaBukuBesar("REMARKS")%></td>
	</tr>
<%

		rsAnalisaBukuBesar.MoveNext
	Wend
end if
%>
	<!--
	<tr bgcolor="#F4F4A0">
		<td colspan="2" align="left"><strong>Total</strong></td>
		<td align="right"><strong></strong></td>		
		<td align="right"><strong></strong></td>	
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong></strong></td>		
		<td align="right"><strong></strong></td>	
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
		<td align="right"><strong>0</strong></td>
	</tr>
	-->

</table>
   
</div>
</body>
</html>