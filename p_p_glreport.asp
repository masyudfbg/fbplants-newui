<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>


<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<%
Dim rsJob, strJob,ACCTID
Dim ScriptName, v_txtfrom, v_txtto, ZeroValue, v_txtJobCode
Dim NO, cClass, nSubTotal, nGrandTotal
DIM v_txtYear, v_txtMonth,sqldetail2,RSDETAIL2,v_txtFcba
Dim n_SaldoSub : n_SaldoSub = 0
Dim n_SaldoGrand : n_SaldoGrand = 0
Dim n_SaldoRow : n_SaldoRow = 0
DIM subtotaldebet,subtotalkredit,saldosubtotal

v_txtfrom	=	Request("glfrom")
v_txtto		=	Request("glfrom")
v_txtYear			= Request("tahun")
v_txtMonth			= Request("bulan")
v_txtJobCode=	Request("coa")
v_txtFcba			= Session("FCBA")

ScriptName	=	Request.ServerVariables("SCRIPT_NAME")
ZeroValue	=	0
nGrandTotal	=	0

if v_txtJobCode = "" then
	strJob		= 	" SELECT FCCODE,FCCODE || ' - ' || FCNAME AS FCNAME FROM JOB WHERE FCBA = '" & Session("FCBA") & "'"
	strJob		=	strJob & " ORDER BY FCCODE ASC"
else 
	strJob		= 	" SELECT FCCODE, FCCODE || ' - ' || FCNAME AS FCNAME FROM JOB WHERE FCBA = '" & Session("FCBA") & "' AND FCCODE = '" & v_txtJobCode & "'"
	strJob		=	strJob & " ORDER BY FCCODE ASC"
end if


Set rsJob = Server.CreateObject("ADODB.Recordset")
rsJob.Open strJob, connstr, adOpenStatic 




%>
<html>

<head>
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>iPS Reporting - Buku Besar</title>
<script language="JavaScript" src="/js/date-picker.js"></script> 
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
<script language="JavaScript" src='/js/x.js'></script>
<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
<link href="../plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="../plugins/chosen/chosen.min.css" rel="stylesheet" />

<link rel="stylesheet" href="../plugins/responsivegridsystem/col.css" media="all">
<link rel="stylesheet" href="../plugins/responsivegridsystem/12cols.css" media="all">
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
</head>

<style type="text/css">
@media print
{
.printhide{display:none;}
}
</style>

<body>
	<div id="wrapper">
		<div id="header-wrapper">
			<div id="header">
				<h4>PT. <%=Session("COMPANY")%></h4>
				<h5><%=Session("BA")%></h5>
			</div>
			<div id="subheader" class="centered">
				<h2>BUKU BESAR</h3>
				<h5>PERIODE&nbsp;:&nbsp;<%=(v_txtYear)%>&nbsp; s.d <%=MonthName(v_txtMonth)%></h5>
			</div>
		</div>
		<br>


<table border="0" width="100%" id="product-table" class="xthf-site">
	<tr>
		<th rowspan=2 >ACCTID</th>
		<th rowspan=2>TANGGAL</th>
		<th rowspan=2>NO. BUKTI</th>
		<th rowspan=2>NO. JURNAL</th>
		<th rowspan=2>KETERANGAN</th>
		<th colspan=3>JUMLAH (Rp.)</th>
	</tr>
	<tr>
		<th colspan= style="color:#FFFFFF">Debet</th>
		<th style="color:#FFFFFF">Credit</th>
		<th style="color:#FFFFFF">Saldo</th>
	</tr>
<tbody>
	<%
		
		While Not rsJob.EOF 
			dim kodeakun,namaakun 
			kodeakun = rsJob("FCCODE")
			namaakun = rsJob("FCNAME")
		
	%>


			<%
					Dim oSQL, oRS
					Dim n_Opening : n_Opening = 0

					oSQL = " Select opening_balance opening" & _
						" from FAC_TRIALBALANCE" & _
						" where fcba = '" & v_txtFcba & "'"

						oSQL = oSQL + "       and FCCODE = '" & rsJob("FCCODE") & "'" 
					If Len(Trim(v_txtYear)) > 0 then
						oSQL = oSQL + "       and period_year = " & v_txtYear & ""
					End If
					If Len(Trim(v_txtMonth)) > 0 then
						oSQL = oSQL + "       and period_month = " & v_txtMonth & ""
					End If

						oSQL = oSQL + "  Order by FCCODE "

					'response.write oSQL 

					Set oRS = Server.CreateObject("ADODB.Recordset")
					oRS.Open oSQL, connstr, adOpenStatic


					If Not oRS.EOF then
						n_Opening = Cdbl(oRS("opening")) 
					Else
						n_Opening = 0 
					End If

					oRS.Close
					set oRS = Nothing

					'n_SaldoSub = n_SaldoSub + n_Opening
					'n_SaldoGrand = n_SaldoGrand + n_Opening
					'n_SaldoRow = n_SaldoRow + n_Opening
			%>


		<tr bgcolor="#9CB770">
			<td align="left" colspan="5" style="padding-left:10px;"><strong>Saldo Awal</strong></td>
			<td align="Right"><strong><%=FormatNumber(0, 0)%></strong></td>
			<td align="Right"><strong><%=FormatNumber(0, 0)%></strong></td>
			<td align="Right"><strong><%=FormatNumber(n_Opening, 0)%></strong></td>
		</tr>
		
		
	
	


	<%	

			sqldetail2 = "Select G.DOCNO, G.JOURNALNO, G.JOURNALDATE, G.COA,C.FCCODE || ' - ' || C.FCNAME COANAME, G.OPENING_BALANCE, " & _
			" G.DEBET, G.CREDIT, G.ENDING_BALANCE, G.CREATEDATE, G.LASTUPDATE, " & _
			" G.FCENTRY, G.FCEDIT, G.FCIP, G.FCBA " & _
			" From   GLEDGERDTL G " & _
			" INNER JOIN JOB C ON (G.COA=C.FCCODE AND G.FCBA=C.FCBA) " & _
			" WHERE TO_NUMBER(TO_CHAR(G.JOURNALDATE,'YYYY'))='" & v_txtYear & "' AND TO_NUMBER(TO_CHAR(G.JOURNALDATE,'MM'))='" & v_txtMonth & "' AND G.FCBA = '" & Session("FCBA") & "' AND G.COA='" & kodeakun & "'" 
		
		Set RSDETAIL2 = Server.CreateObject("ADODB.Recordset")
		RSDETAIL2.Open sqldetail2, connstr, adOpenStatic 

		subtotaldebet = 0
		subtotalkredit= 0
		saldosubtotal = 0
		While Not RSDETAIL2.EOF 

		'saldosubtotal =n_Opening + RSDETAIL2("DEBET") - RSDETAIL2("CREDIT")
		'saldosubtotal =n_Opening +  RSDETAIL2("CREDIT") - RSDETAIL2("DEBET")
	
	
	%>

	<tr>
		<td>
			<%=RSDETAIL2("COANAME")%>
		</td>
		<td colspan=>
			<%=RSDETAIL2("JOURNALDATE")%>
		</td>
		<td>
			<%=RSDETAIL2("DOCNO")%>
		</td>
		<td>
			<%=RSDETAIL2("JOURNALNO")%>
		</td>
		<td>
			
		</td>
		<td align="right">
			<%=formatnumber(RSDETAIL2("DEBET"),0)%>
		</td>
		<td align="right">
			<%=formatnumber(RSDETAIL2("CREDIT"),0)%>
		</td>
		<td align="right"></td>
	</tr>
	<% 

		subtotaldebet = subtotaldebet + RSDETAIL2("DEBET")
		subtotalkredit= subtotalkredit + RSDETAIL2("CREDIT")
		if subtotaldebet = 0 then 
			saldosubtotal = cdbl(subtotalkredit) - cdbl(subtotaldebet) + cdbl(n_Opening)
		else 
			saldosubtotal = cdbl(n_Opening)+cdbl(subtotaldebet)-cdbl(subtotalkredit)
		end if
		RSDETAIL2.MoveNext
		wend

		
	%>

	<tr bgcolor="#9CB770">
			<td align="left" colspan="5" style="padding-left:10px;"><strong>Saldo Akhir</strong></td>
			<td align="Right"><strong><%=FormatNumber(subtotaldebet, 0)%></strong></td>
			<td align="Right"><strong><%=FormatNumber(subtotalkredit, 0)%></strong></td>
			<td align="Right"><strong><%=FormatNumber(saldosubtotal, 0)%></strong></td>
			
	</tr>
	
	<% 
		rsJob.MoveNext
		wend

		
	%>

</tbody>
</table>
</div>
</body>
</html>
<!--#include file="Bottom_Report.asp"-->