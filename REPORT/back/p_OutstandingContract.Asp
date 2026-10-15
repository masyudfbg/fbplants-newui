<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, v_txtContractorCode, v_txtNoSPK
Dim ScriptName
ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
v_txtContractorCode=	trim(request("txtContractorCode"))
v_txtNoSPK=	trim(request("txtNoSPK"))
%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPla'S Reporting - Outstanding Kontrak</title>
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

				<td><font face="verdana" size="1">KONTRAKTOR</font></td><td><select name="txtContractorCode" ><option value=""></option>
				<%
				Dim rsContractor, strContractor
				strContractor	=	" SELECT FCCODE, FCNAME FROM CONTRACTOR WHERE ACTIVATION = 'Y' AND FCBA = '" & Session("FCBA") & "' ORDER BY FCCODE ASC"
				
				Set rsContractor = Server.CreateObject("ADODB.Recordset")
				rsContractor.Open strContractor, connstr, adOpenStatic 
				
				if not rsContractor.eof then
				Do While Not rsContractor.eof
				if rsContractor("FCCODE") = v_txtContractorCode then%>
						<option value="<%=rsContractor("FCCODE")%>" Selected><%=rsContractor("FCCODE")%> - <%=rsContractor("FCNAME")%></option> 
					 <%Else%>
						<option value="<%=rsContractor("FCCODE")%>"><%=rsContractor("FCCODE")%> - <%=rsContractor("FCNAME")%></option>  
					<%End if
					rsContractor.movenext
				loop
				end if
				rsContractor.Close
				set rsContractor = Nothing
				%>
				</select></td>
				<td><font face="verdana" size="1">NO. SPK</font></td>
				<td><input type="text" maxlength="50" name="txtNoSPK" size="30" value="<%=v_txtNoSPK%>" title="Input No SPK"></td>
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
		<h3>LAPORAN OUTSTANDING KONTRAK</h3>
		<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(mm)%>&nbsp;<%=yyyy%></h4>
	</div>
	</div>
<table border="0" width="100%" id="product-table">
	<tr>
		<th style="color:#FFFFFF" rowspan="3">NO. SPK</th>
		<th style="color:#FFFFFF" rowspan="3">JATUH TEMPO</th>
		<th style="color:#FFFFFF" rowspan="3">KODE KONTRAKTOR</th>
		<th style="color:#FFFFFF" rowspan="3">NAMA KONTRAKTOR</th>
		<th style="color:#FFFFFF" rowspan="3">KODE PEKERJAAN</th>
		<th style="color:#FFFFFF" rowspan="3">JENIS PEKERJAAN</th>
		<th style="color:#FFFFFF" colspan="4">QUANTITY</th>
		<th style="color:#FFFFFF" rowspan="3">KETERANGAN</th>
	</tr>
	<tr>
		<th style="color:#FFFFFF" rowspan="2">KONTRAK</th>
		<th style="color:#FFFFFF" colspan="2">WIP</th>
		<th style="color:#FFFFFF" rowspan="2">SISA</th>		
	</tr>
	<tr>
		<th style="color:#FFFFFF">BI</th>
		<th style="color:#FFFFFF">SBI</th>	
	</tr>
<%
Dim rsOutstandingSPK, strOutstandingSPK, rowno
rowno		=	0
strOutstandingSPK		=	" SELECT theData.agreementcode," & _
							" theData.lineno," & _
							" theMaster.contractorcode," & _
							" CONTRACTOR.FCNAME contractorname," & _
							" theMaster.job," & _
							" JOB.FCNAME jobname," & _
							" theMaster.duedate," & _
							" theData.contract_qty," & _
							" theData.bi_qty," & _
							" theData.sbi_qty," & _
							" (theData.contract_qty - theData.sbi_qty) remain_qty," & _
							" theMaster.remarks" & _
							" FROM ( SELECT agreementcode," & _
							" lineno," & _
							" SUM (contract_qty) contract_qty," & _
							" SUM (bi_qty) bi_qty," & _
							" SUM (sbi_qty) sbi_qty," & _
							" fcba" & _
							" FROM (SELECT a.agreementcode," & _
							" a.lineno," & _
							" a.qty contract_qty," & _
							" 0 bi_qty," & _
							" 0 sbi_qty," & _
							" a.fcba" & _
							" FROM AGREEMENTDETAIL a" & _
							" WHERE a.fcba = '" & Session("FCBA") & "'" & _
							" UNION" & _
							" SELECT a.agreementcode," & _
							" a.agreementlineno," & _
							" 0 contract_qty," & _
							" SUM (a.qty) bi_qty," & _
							" 0 sbi_qty," & _
							" a.fcba" & _
							" FROM workinprogress a" & _
							" WHERE a.fcba = '" & Session("FCBA") & "'" & _
							" AND TO_NUMBER (TO_CHAR (wipdate, 'YYYY')) =" & yyyy & _
							" AND TO_NUMBER (TO_CHAR (wipdate, 'MM')) =" & MM & _
							" GROUP BY a.agreementcode, a.agreementlineno, a.fcba" & _
							" UNION" & _
							" SELECT a.agreementcode," & _
							" a.agreementlineno," & _
							" 0 contract_qty," & _
							" 0 bi_qty," & _
							" SUM (a.qty) sbi_qty," & _
							" a.fcba" & _
							" FROM workinprogress a" & _
							" WHERE a.fcba = '" & Session("FCBA") & "'" & _
							" AND TO_NUMBER (TO_CHAR (wipdate, 'YYYY')) <=" & yyyy & _
							" AND TO_NUMBER (TO_CHAR (wipdate, 'MM')) <=" & MM & _
							" GROUP BY a.agreementcode, a.agreementlineno, a.fcba)" & _
							" GROUP BY agreementcode, lineno, fcba) theData," & _
							" agreementdetail theMaster," & _
							" job," & _
							" contractor" & _
							" WHERE theData.agreementcode = theMaster.agreementcode" & _
							" AND theData.lineno = theMaster.lineno" & _
							" AND theMaster.contractorcode = contractor.fccode" & _
							" AND theMaster.job = job.fccode" & _
							" AND theData.fcba = theMaster.fcba" & _
							" AND theMaster.fcba = job.fcba" & _
							" AND theMaster.fcba = contractor.fcba"

if  LEN(TRIM(v_txtContractorCode)) > 0 then
	strOutstandingSPK	=	strOutstandingSPK & " AND theMaster.contractorcode  = '" & v_txtContractorCode & "'"
end if
if  LEN(TRIM(v_txtNoSPK)) > 0 then
	strOutstandingSPK	=	strOutstandingSPK & " AND theData.agreementcode  = '" & v_txtNoSPK & "'"
end if

Set rsOutstandingSPK = Server.CreateObject("ADODB.Recordset")
rsOutstandingSPK.Open strOutstandingSPK, connstr, adOpenStatic
if Not rsOutstandingSPK.eof then

	While Not rsOutstandingSPK.eof
			rowno	=	rowno	+	1
%>
	<tr>
		<td align="left"><%=rsOutstandingSPK("agreementcode")%></td>
		<td align="right"><%=FbpNum2(rsOutstandingSPK("duedate"))%></td>
		<td align="center"><%=rsOutstandingSPK("contractorcode")%></td>
		<td align="left"><%=rsOutstandingSPK("contractorname")%></td>
		<td align="center"><%=rsOutstandingSPK("job")%></td>
		<td align="left"><%=rsOutstandingSPK("jobname")%></td>
		<td align="right"><%=FormatNumber(rsOutstandingSPK("contract_qty"),2)%></td>
		<td align="right"><%=FormatNumber(rsOutstandingSPK("bi_qty"),2)%></td>
		<td align="right"><%=FormatNumber(rsOutstandingSPK("sbi_qty"),2)%></td>
		<td align="right"><%=FormatNumber(rsOutstandingSPK("remain_qty"),2)%></td>
		<td align="left"><%=rsOutstandingSPK("REMARKS")%></td>
	</tr>
<%

		rsOutstandingSPK.MoveNext
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