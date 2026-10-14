<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
dIM v_txtYear, v_txtMonth
Dim MM, yyyy, mon, Yr, MinYr, MaxYr, i, v_txtContractorCode, v_txtNoSPK
Dim ScriptName


v_txtYear	=	Request("txtYear")
v_txtMonth	=	Request("txtMonth")

If Len(Trim(v_txtYear)) = 0 Then
	v_txtYear = Year(Session("v_Costbook"))
End If

If Len(Trim(v_txtMonth)) = 0 Then
	v_txtMonth = Month(Session("v_Costbook"))
End If

ScriptName =	Request.ServerVariables("SCRIPT_NAME")
v_txtContractorCode = trim(request("txtContractorCode"))
v_txtNoSPK = trim(request("txtNoSPK"))
%>
<html>

<head>
<title>ePlaS Reporting - Outstanding Kontrak</title>
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
	<table id="form-table" class="centered-margin">
		<form name="frmBTL" method="post" action="<%=ScriptName%>">
			<tr>
				<td><font face="verdana" size="1">TAHUN</font></td>
				<td>
<% 
    mon = Month(Session("v_Costbook"))
    If len(MM)=0 then MM=mon

    Yr = Year(Session("v_Costbook"))
    MinYr = 2015
    MaxYr = Yr  + 2
    If len(yyyy)= 0 then yyyy = Yr
%>
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
				<td>
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
<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
	<tr>
    <th colspan="50" style="background-color:#FFFFFF; border:none;">  
	<div id="header-wrapper">
		<div id="head_judul" style="float:left; text-align:left;">
		<h4>PT. <%=Session("COMPANY")%></h4>
		<h5><%=Session("BA")%></h5>
	</div>
	<div id="subheader" class="centered" style="width:50%; margin:auto;">
		<h2>LAPORAN OUTSTANDING KONTRAK</h2>
		<h4>PERIODE&nbsp;:&nbsp;<%=Bulan(v_txtMonth)%>&nbsp;<%=v_txtYear%></h4>
	</div>
	</div>
       </th>
  </tr>
  
	<tr>
			<th style="color:#FFFFFF" rowspan="3">NO. SPK</th>
			<th style="color:#FFFFFF" rowspan="3">JATUH TEMPO</th>
			<th style="color:#FFFFFF" rowspan="3">NAMA KONTRAKTOR</th>
			<th style="color:#FFFFFF" rowspan="3">JENIS PEKERJAAN</th>
			<th style="color:#FFFFFF" rowspan="2" colspan="2">KONTRAK</th>
			<th style="color:#FFFFFF" colspan="4">PROGRESS</th>
			<th style="color:#FFFFFF" rowspan="2" colspan="2">SISA</th>
		</tr>
		<tr>
			<th style="color:#FFFFFF" colspan="2">BULAN INI</th>
			<th style="color:#FFFFFF" colspan="2">S.D BULAN INI</th>	
		</tr>
		<tr>
			<th style="color:#FFFFFF">QTY</th>
			<th style="color:#FFFFFF">AMOUNT</th>
			<th style="color:#FFFFFF">QTY</th>
			<th style="color:#FFFFFF">AMOUNT</th>
			<th style="color:#FFFFFF">QTY</th>
			<th style="color:#FFFFFF">AMOUNT</th>
			<th style="color:#FFFFFF">QTY</th>
			<th style="color:#FFFFFF">AMOUNT</th>		
		</tr>
	</thead>
	<tbody>
<%
Dim rsOutstandingSPK, strOutstandingSPK, rowno
rowno		=	0
' strOutstandingSPK		=	" SELECT theData.agreementcode," & _
							' " theData.lineno," & _
							' " theMaster.contractorcode," & _
							' " CONTRACTOR.FCNAME contractorname," & _
							' " theMaster.job," & _
							' " JOB.FCNAME jobname," & _
							' " theMaster.duedate," & _
							' " theData.contract_qty," & _
							' " theData.bi_qty," & _
							' " theData.sbi_qty," & _
							' " (theData.contract_qty - theData.sbi_qty) remain_qty," & _
							' " theMaster.remarks" & _
							' " FROM ( SELECT agreementcode," & _
							' " lineno," & _
							' " SUM (contract_qty) contract_qty," & _
							' " SUM (bi_qty) bi_qty," & _
							' " SUM (sbi_qty) sbi_qty," & _
							' " fcba" & _
							' " FROM (SELECT a.agreementcode," & _
							' " a.lineno," & _
							' " a.qty contract_qty," & _
							' " 0 bi_qty," & _
							' " 0 sbi_qty," & _
							' " a.fcba" & _
							' " FROM AGREEMENTDETAIL a" & _
							' " WHERE a.fcba = '" & Session("FCBA") & "'" & _
							' " UNION" & _
							' " SELECT a.agreementcode," & _
							' " a.agreementlineno," & _
							' " 0 contract_qty," & _
							' " SUM (a.qty) bi_qty," & _
							' " 0 sbi_qty," & _
							' " a.fcba" & _
							' " FROM workinprogress a" & _
							' " WHERE a.fcba = '" & Session("FCBA") & "'" & _
							' " AND TO_NUMBER (TO_CHAR (wipdate, 'YYYY')) =" & yyyy & _
							' " AND TO_NUMBER (TO_CHAR (wipdate, 'MM')) =" & MM & _
							' " GROUP BY a.agreementcode, a.agreementlineno, a.fcba" & _
							' " UNION" & _
							' " SELECT a.agreementcode," & _
							' " a.agreementlineno," & _
							' " 0 contract_qty," & _
							' " 0 bi_qty," & _
							' " SUM (a.qty) sbi_qty," & _
							' " a.fcba" & _
							' " FROM workinprogress a" & _
							' " WHERE a.fcba = '" & Session("FCBA") & "'" & _
							' " AND TO_NUMBER (TO_CHAR (wipdate, 'YYYY')) <=" & yyyy & _
							' " AND TO_NUMBER (TO_CHAR (wipdate, 'MM')) <=" & MM & _
							' " GROUP BY a.agreementcode, a.agreementlineno, a.fcba)" & _
							' " GROUP BY agreementcode, lineno, fcba) theData," & _
							' " agreementdetail theMaster," & _
							' " job," & _
							' " contractor" & _
							' " WHERE theData.agreementcode = theMaster.agreementcode" & _
							' " AND theData.lineno = theMaster.lineno" & _
							' " AND theMaster.contractorcode = contractor.fccode" & _
							' " AND theMaster.job = job.fccode" & _
							' " AND theData.fcba = theMaster.fcba" & _
							' " AND theMaster.fcba = job.fcba" & _
							' " AND theMaster.fcba = contractor.fcba"

' if  LEN(TRIM(v_txtContractorCode)) > 0 then
	' strOutstandingSPK	=	strOutstandingSPK & " AND theMaster.contractorcode  = '" & v_txtContractorCode & "'"
' end if
' if  LEN(TRIM(v_txtNoSPK)) > 0 then
	' strOutstandingSPK	=	strOutstandingSPK & " AND theData.agreementcode  = '" & v_txtNoSPK & "'"
' end if

strOutstandingSPK	=	"   SELECT fcba," & _
						"          agreementcode," & _
						"          finishdate," & _
						"          contractorcode," & _
						"          contractorname," & _
						"          jobcode," & _
						"          jobname," & _
						"          rate," & _
						"          contract_qty," & _
						"          qty_final_bi," & _
						"          qty_final_sbi," & _
						"          amount_contract," & _
						"          amount_final_bi," & _
						"          amount_final_sbi," & _
						"          qty_sisa," & _
						"          amount_sisa" & _
						" FROM ( " & _
						"   SELECT ca.fcba," & _
						"          ca.agreementcode," & _
						"          ca.finishdate," & _
						"          ca.contractorcode," & _
						"          CONTRACTOR_API.GET_CONTRACTORNAME (contractorcode, ca.fcba)" & _
						"             contractorname," & _
						"          MAX (cad.jobcode) jobcode," & _
						"          MAX (JOB_API.GET_JOBCODE_NAME (cad.jobcode, ca.fcba)) jobname," & _
						"          NVL (SUM (cad.rate), 0) rate," & _
						"          NVL (SUM (cad.qty), 0) contract_qty," & _
						"          NVL (SUM (wf.qty_bi), 0) qty_final_bi," & _
						"          NVL (SUM (wf.qty_sbi), 0) qty_final_sbi," & _
						"          NVL (SUM (cad.rate) * SUM (cad.qty), 0) amount_contract," & _
						"          NVL (SUM (cad.rate) * SUM (wf.qty_bi), 0) amount_final_bi," & _
						"          NVL (SUM (cad.rate) * SUM (wf.qty_sbi), 0) amount_final_sbi," & _
						"          NVL (SUM (wf.qty_bi) - SUM (wf.qty_sbi), 0) qty_sisa," & _
						"          NVL (" & _
						"               (SUM (cad.rate) * SUM (wf.qty_bi))" & _
						"             - (SUM (cad.rate) * SUM (wf.qty_sbi))," & _
						"             0)" & _
						"             amount_sisa" & _
						"     FROM contractagreementdetail cad" & _
						"          INNER JOIN contractagreement ca" & _
						"             ON (ca.agreementcode = cad.agreementcode AND ca.fcba = cad.fcba)" & _
						"          LEFT JOIN" & _
						"          (  SELECT wf.fcba," & _
						"                    wf.agreementcode," & _
						"                    wf.agreementlineno," & _
						"                    NVL (SUM (wf.qty), 0) qty_bi," & _
						"                    0 qty_sbi" & _
						"               FROM workinprogressfinal wf" & _
						"              WHERE     fcba = '"& session("fcba") &"'" & _
						"                    AND TO_NUMBER (TO_CHAR (wipdate, 'yyyy')) = '"& v_txtYear &"'" & _
						"                    AND TO_NUMBER (TO_CHAR (wipdate, 'mm')) = '"& v_txtMonth &"'" & _
						"           GROUP BY wf.fcba, wf.agreementcode, wf.agreementlineno" & _
						"           UNION ALL" & _
						"             SELECT wf.fcba," & _
						"                    wf.agreementcode," & _
						"                    wf.agreementlineno," & _
						"                    0 qty_bi," & _
						"                    NVL (SUM (wf.qty), 0) qty_sbi" & _
						"               FROM workinprogressfinal wf" & _
						"              WHERE     fcba = '"& session("fcba") &"'" & _
						"                    AND TO_NUMBER (TO_CHAR (wipdate, 'yyyy')) = '"& v_txtYear &"'" & _
						"                    AND TO_NUMBER (TO_CHAR (wipdate, 'mm')) <= '"& v_txtMonth &"'" & _
						"           GROUP BY wf.fcba, wf.agreementcode, wf.agreementlineno) wf" & _
						"             ON     wf.fcba = cad.fcba" & _
						"                AND wf.agreementcode = cad.agreementcode" & _
						"                AND cad.lineno = wf.agreementlineno" & _
						"    WHERE ca.fcba = '"& session("fcba") &"'"  & _
						" GROUP BY ca.fcba," & _
						"          ca.agreementcode," & _
						"          ca.finishdate," & _
						"          ca.contractorcode," & _ 
						"          CONTRACTOR_API.GET_CONTRACTORNAME (contractorcode, ca.fcba) ) WHERE  qty_sisa > 0 AND amount_sisa > 0" 
						
						
						
If  Len(Trim(v_txtContractorCode)) > 0 then
	strOutstandingSPK	=	strOutstandingSPK & " AND CONTRACTOR_API.GET_CONTRACTORNAME (contractorcode, fcba)  = '" & v_txtContractorCode & "'"
End If
If  Len(Trim(v_txtNoSPK)) > 0 then
	strOutstandingSPK	=	strOutstandingSPK & " AND ca.agreementcode  = '" & v_txtNoSPK & "'"
End If

	' strOutstandingSPK	=	strOutstandingSPK & " GROUP BY ca.fcba," & _
												' "          ca.agreementcode," & _
												' "          ca.finishdate," & _
												' "          ca.contractorcode," & _ 
												' "          CONTRACTOR_API.GET_CONTRACTORNAME (contractorcode, ca.fcba)"  

	strOutstandingSPK	=	strOutstandingSPK & " ORDER BY contractorcode, agreementcode " 


' Response.Write strOutstandingSPK

Set rsOutstandingSPK = Server.CreateObject("ADODB.Recordset")
rsOutstandingSPK.Open strOutstandingSPK, connstr, adOpenStatic

If rsOutstandingSPK.BOF and rsOutstandingSPK.EOF then
%>
		<tr>
			<td align="left" colspan="12">NO DATA AVAILABLE</td>
		</tr>
<%
Else
	While Not rsOutstandingSPK.eof
			rowno	=	rowno	+	1
%>
	
		<tr>
			<td align="left"><%=rsOutstandingSPK("agreementcode")%></td>
			<td align="left"><%=rsOutstandingSPK("finishdate")%></td>
			<td align="left"><%=rsOutstandingSPK("contractorcode")%> : <%=rsOutstandingSPK("contractorname")%></td>
			<td align="left"><%=rsOutstandingSPK("JOBNAME")%></td>
			<td align="right"><%=FormatNumber(rsOutstandingSPK("contract_qty"),0)%></td>
			<td align="right"><%=FormatNumber(rsOutstandingSPK("amount_contract"),0)%></td>
			<td align="right"><%=FormatNumber(rsOutstandingSPK("qty_final_bi"),0)%></td>
			<td align="right"><%=FormatNumber(rsOutstandingSPK("amount_final_bi"),0)%></td>
			<td align="right"><%=FormatNumber(rsOutstandingSPK("qty_final_sbi"),0)%></td>
			<td align="right"><%=FormatNumber(rsOutstandingSPK("amount_final_sbi"),0)%></td>
			<td align="right"><%=FormatNumber(rsOutstandingSPK("QTY_SISA"),0)%></td>
			<td align="right"><%=FormatNumber(rsOutstandingSPK("AMOUNT_SISA"),0)%></td>
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
	</tbody>
</table>
   
</div>
</body>
</html>