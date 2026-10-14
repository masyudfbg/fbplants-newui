<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<html>

<head>
<title>Reporting - Nota Debit Unit</title>
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

<style type="text/css">
@media print
{
body{font-family:'Arial Narrow';font-size:10px;}
.printhide{display:none;}
}
.hidden{
  /* Transparent text, should work on any background colour */
  color: rgba(0,0,0,0); 
  font-size: 0;
}
.ratakanan{
	text-align:right;
}
.huruftebal {
					font-weight: bold;
			}
</style>
</head>

	 <%

	dim region,company,bulan,tahun,costelemen,t1,r1
	dim txtcari,fcba,bkmfrom,bkmto,idview,actfrom,actto,sectioncode,statustm
	dim filterdiv,filtersource,filterfcba,filterprd,filterprdlalu,filtercostelemen,filterprdsbi,filtertmtbm
	dim prev_bulanfrom,prev_bulanto
	dim bulansplitfrom,tglsplitfrom,thnsplitfrom,tglfrom,tglto,tglawalsbi
	dim bulansplitto,tglsplitto,thnsplitto,xbln1,xthn1
	dim rphablnlalu
	dim	fisikblnlaluttl,amountblnlaluttl,rphablnlaluttl,rkb_outputblnlaluttl,rkb_amountblnlaluttl,rkbhablnlaluttl
	dim	fisikblninittl,amountblninittl,rphablninittl,rkb_outputblninittl,rkb_amountblninittl,rkbhablninittl
	dim fisiksbittl,amountsbittl,rphasbittl,rkb_outputsbittl,rkb_amountsbittl,rkbhasbittl
	dim sqlexec,periode,sqlpt,rspt,filtercompany1,fcbaloop,companyview

	company 	= Request("company")
	fcbaFrom 	= Request("fcbaFrom")
	fcbaTo 		= Request("fcbaTo")
	bulan   	= Request("bulan")
	tahun   	= Request("tahun")
	
	if fcbaFrom = "null" or fcbafrom="-" then 
		fcbaFrom = ""
	end if
	
	if fcbaTo = "null" or fcbaTo="-" then 
		fcbaTo = ""
	end if
	

	q = "Select CompanyName from company where companycode='" & company & "' "
		set rsCat = connStr.Execute(q)
		if not rsCat.eof then
			CompanyName=rsCat("CompanyName")
		end if
	    rsCat.close()
		rsCat	=	null

    sqlview = 	" WITH x AS ( " & _
				" SELECT c.fccompanycode, b.sourcecode, b.docno, b.fcba, b.docdate, b.accountingperiod, b.accountingyear," & _
				" 	   a.description, amountdebit, chargeunit, chargejob, jobmaster_api.get_fcname(chargejob) ChargejobName ,chargecode," & _
				" 	   ROW_NUMBER() OVER (PARTITION BY REPLACE(b.docno, '-INT', ''),b.fcba,chargeunit ORDER BY b.docno) AS rw" & _
				" FROM bkmjournaldtl a" & _
				" INNER JOIN bkmjournalhdr b ON (a.docno = b.docno AND a.fcba = b.fcba)" & _
				" INNER JOIN businessunit c ON (a.fcba = c.fccode)" & _
				" WHERE a.jobcode = '1114001'" & _
				" AND b.accountingyear = " & tahun & " AND b.accountingperiod = " & bulan & " AND c.fccompanycode = '" & company & "'" & _
				" AND sourcecode in ('FAC-BNK','FAC-VRA','FAC-WSI','OPR-BKM','OPR-WIP','FAC-JRV','MAM-STT')" & _
				" AND amountdebit <> 0" & _
				" )," & _
				" v AS" & _
				" ( " & _
				" SELECT c.fccompanycode, b.sourcecode, b.docno,b.docdate, REPLACE(b.docno, '-INT', '') AS linkdocno, " & _
				"		 b.fcba, b.accountingperiod, b.accountingyear, a.description, amountdebit*-1 AS amountkredit, " & _ 
				"		 b.fcba chargeunit, jobcode chargejob, jobmaster_api.get_fcname(jobcode) ChargejobName ,locationcode chargecode, " & _
				" 		 ROW_NUMBER() OVER (PARTITION BY REPLACE(b.docno, '-INT', '') ORDER BY b.docno) AS rn " & _
				" FROM bkmjournaldtl a INNER JOIN bkmjournalhdr b ON (a.docno = b.docno AND a.fcba = b.fcba) INNER JOIN businessunit c ON (a.fcba = c.fccode) " & _
				" WHERE " & _
				" a.docno LIKE '%-INT'" & _
				" AND b.accountingyear = " & tahun & " AND b.accountingperiod = " & bulan & " AND c.fccompanycode = '" & company & "' " & _ 
				" AND sourcecode in ('FAC-BNK','OPR-BKM','OPR-WIP','FAC-VRA','FAC-WSI') AND amountdebit <> 0 " & _
				" )," & _
				" w AS " & _
				" 	( " & _
				" 	SELECT c.fccompanycode, b.sourcecode, b.docno,b.docdate, a.bkmno AS linkdocno, b.fcba, " & _
				" 	b.accountingperiod, b.accountingyear, a.description, amountdebit*-1 AS amountkredit, " & _
				" 	b.fcba chargeunit, jobcode chargejob, jobmaster_api.get_fcname(jobcode) ChargejobName ,locationcode chargecode, " & _
				" 	ROW_NUMBER() OVER (PARTITION BY a.bkmno ORDER BY b.docno) AS rn " & _
				" 	FROM bkmjournaldtl a " & _
				" 	INNER JOIN bkmjournalhdr b ON (a.docno = b.docno AND a.fcba = b.fcba) " & _
				" 	INNER JOIN businessunit c ON (a.fcba = c.fccode) " & _
				" 	WHERE " & _
				" 	b.accountingyear = " & tahun & " AND b.accountingperiod = " & bulan & " AND c.fccompanycode = '" & company & "' " & _
				" 	AND sourcecode in ('MAM-SRT') AND amountdebit <> 0 " & _
				" 	), " & _
				" 	z AS " & _
				" 	( " & _
				" 		SELECT c.fccompanycode, b.sourcecode, b.docno,b.docdate, a.bkmno AS linkdocno, b.fcba, " & _
				" 		b.accountingperiod, b.accountingyear, a.description,  amountkredit, " & _
				" 		chargeunit chargeunit, jobcode chargejob, jobmaster_api.get_fcname(jobcode) ChargejobName ,locationcode chargecode, " & _
				" 		ROW_NUMBER() OVER (PARTITION BY a.bkmno ORDER BY b.docno) AS rn " & _
				" 		FROM bkmjournaldtl a " & _
				" 		INNER JOIN bkmjournalhdr b ON (a.docno = b.docno AND a.fcba = b.fcba) " & _
				" 		INNER JOIN businessunit c ON (a.fcba = c.fccode) " & _
				" 		WHERE  " & _
				" 		b.accountingyear = " & tahun & " AND b.accountingperiod = " & bulan & " AND c.fccompanycode = '" & company & "'" & _
				" 		AND sourcecode in ('FAC-JRV') AND amountkredit <> 0 " & _
				" 		) " & _
				" SELECT x.accountingperiod, x.accountingyear,x.fccompanycode, x.sourcecode, x.docno,x.docdate, x.fcba," & _
				" 	   x.description,  y.chargeunit, y.chargejob||'-'|| y.chargejobName jobcd, y.chargecode,x.amountdebit, " & _
				" 	   Nvl(y.amountkredit,0 )amountkredit, y.docno chargeDocNo,y.docdate chargeDocDate,y.description chargeDescription,x.rw " & _
				" FROM x" & _
				" LEFT JOIN (" & _
				" 	SELECT * FROM v " & _
				" ) y ON (x.fccompanycode = y.fccompanycode AND x.docno = y.linkdocno  AND x.chargeunit=y.fcba AND Abs(amountdebit)=Abs(amountkredit))" & _
				" WHERE x.fcba Like '" & fcbaFrom & "%' and x.chargeunit Like '" & fcbaTo & "%' AND x.sourcecode in ('FAC-BNK','OPR-BKM','OPR-WIP') " & _
				" UNION ALL " & _ 
				" SELECT x.accountingperiod, x.accountingyear,x.fccompanycode, x.sourcecode, x.docno,x.docdate, x.fcba," & _
				" 	   x.description,  y.chargeunit, y.chargejob||'-'|| y.chargejobName jobcd, y.chargecode,x.amountdebit, " & _
				" 	   Nvl(y.amountkredit,0 )amountkredit, y.docno chargeDocNo,y.docdate chargeDocDate,y.description chargeDescription,x.rw " & _
				" FROM x" & _
				" LEFT JOIN (" & _
				" 	SELECT * FROM v WHERE rn = 1" & _
				" ) y ON (x.fccompanycode = y.fccompanycode AND x.docno = y.linkdocno AND Abs(amountdebit)=Abs(amountkredit))" & _
				" WHERE x.fcba Like '" & fcbaFrom & "%' and y.chargeunit Like '" & fcbaTo & "%' AND x.sourcecode in ('FAC-VRA','FAC-WSI') " & _
				" UNION ALL " & _ 
				" SELECT x.accountingperiod, x.accountingyear,x.fccompanycode, x.sourcecode, x.docno,x.docdate, x.fcba, x.description, " & _
				" y.chargeunit, y.chargejob||'-'|| y.chargejobName jobcd, y.chargecode,x.amountdebit, " & _
				" Nvl(y.amountkredit,0 )amountkredit, y.docno chargeDocNo,y.docdate chargeDocDate,y.description chargeDescription,x.rw " & _
				" FROM x LEFT JOIN ( SELECT * FROM w  ) y ON (x.fccompanycode = y.fccompanycode AND x.docno = y.linkdocno AND x.rw=y.rn) " & _
				" WHERE x.fcba Like '" & fcbaFrom & "%' and y.chargeunit Like '" & fcbaTo & "%' " & _
				" AND x.sourcecode in ('MAM-STT','MAM-SRT')" & _
				" UNION ALL " & _ 
				" SELECT x.accountingperiod, x.accountingyear,x.fccompanycode, x.sourcecode, x.docno,x.docdate, x.fcba, x.description, " & _ 
				" x.chargeunit, y.chargejob||'-'|| y.chargejobName jobcd, y.chargecode,x.amountdebit, " & _ 
				" Nvl(y.amountkredit,0 )amountkredit, y.docno chargeDocNo,y.docdate chargeDocDate,y.description chargeDescription,x.rw " & _ 
				" FROM x LEFT JOIN ( SELECT * FROM z ) y ON (x.fccompanycode = y.fccompanycode AND Abs(amountdebit)=Abs(amountkredit) ) " & _ 
				" WHERE x.fcba Like '" & fcbaFrom & "%' " & _
				" AND x.sourcecode in ('FAC-JRV') and y.chargejob = '1114001' " & _ 
				" ORDER BY sourcecode,docdate,docno,rw "

	'Response.Write sqlview

	Set rsdata = Server.CreateObject("ADODB.Recordset")
	rsdata.CursorType = 3
	rsdata.CursorLocation = 3
	rsdata.Open sqlview, connstr
	 
	%>
	<body onload="">
	<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
	 
	<table width="100%" id="product-table" class="xthf-site" id="">
		 <thead>
        	 <tr>
				<th colspan="28" style=" background-color:#FFFFFF; color:#000000; border:none">
					<div id="header-wrapper">
					<div id="head_judul" style="float:left; text-align:left;">
							<h4>PT. <%=CompanyName%></h4>
					</div>
					<div id="subheader" class="centered" style="width:50%; margin:auto;">
							<h1 class="judul">Laporan Nota Debit - Unit</h1>
							<h3>Periode: <%=textMonth(bulan)%> <%=tahun%> </h5>
					</div>
						<div style="clear:both;"></div>
					</div>
				</th>
			</tr>
				<!--<div class="printhide">-->
			
		<tr style="background-color:#00A300;color:#fff;">
			<th>Source</th>
		    <th>No.Doc</th>
			<th>Date.Doc</th>
			<th>Description</th>
			<th>Fcba</th>
			<th>ChargeUnit</th>
			<th>ChargeJob</th>
			<th>ChargeCode</th>
			<th>Db</th>
			<th>Cr</th>
			<th>No.Doc</th>
			<th>Date.Doc</th>
			<th>Description</th>
		</tr>
		</thead>
		<%
		Dim SourceCode,SumDb,SumCr
        SumDb=0
		SumCr=0		
		SourceCode = ""
				While Not rsdata.EOF
			
		%>
		
		<% if SourceCode <> rsdata("sourcecode") then %>
		<tr class=''>
			<td colspan=13 align="Left" style="background-color:#00A300;color:#fff;" ><%=rsdata("sourcecode")%></td>
		</tr>
			<tr class=''>
				<td>&nbsp;</td>
				<td><%=rsdata("docno")%></td>
				<td><%=rsdata("docdate")%></td>
				<td><%=rsdata("description")%></td>
				<td><%=rsdata("fcba")%></td>
				<td><%=rsdata("chargeunit")%></td>
				<td><%=rsdata("jobcd")%></td>
				<td><%=rsdata("chargecode")%></td>
				<td class="ratakanan"><%=FormatNumber(rsdata("amountdebit"))%></td>
				<td class="ratakanan"><%=FormatNumber(rsdata("amountkredit"))%></td>
				<td><%=rsdata("chargeDocNo")%></td>
				<td><%=rsdata("chargeDocDate")%></td>
				<td><%=rsdata("chargeDescription")%></td>
			</tr>
		<% else %>
		<tr class=''>
			<td>&nbsp;</td>
			<td><%=rsdata("docno")%></td>
			<td><%=rsdata("docdate")%></td>
			<td><%=rsdata("description")%></td>
			<td><%=rsdata("fcba")%></td>
			<td><%=rsdata("chargeunit")%></td>
			<td><%=rsdata("jobcd")%></td>
			<td><%=rsdata("chargecode")%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("amountdebit"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("amountkredit"))%></td>
			<td><%=rsdata("chargeDocNo")%></td>
			<td><%=rsdata("chargeDocDate")%></td>
			<td><%=rsdata("chargeDescription")%></td>
		</tr>
		<%
			end if
			SourceCode = rsdata("sourcecode")
			SumDb = SumDb + cdbl(rsdata("amountdebit"))
			SumCr = SumCr + cdbl(rsdata("amountkredit"))
			
			rsdata.MoveNext
			Wend
			rsdata.close()
			rsdata=	null
		%>
    <tr bgcolor="#F4F4A0">
		<td align="Center" colspan="8"><font face="Verdana" size="1"><b>T o t a l</b></font></td>
		<td align="Right" > <font face="Verdana" size="1"><b><%=FormatNumber(SumDb,2)%></b></font></td>
		<td align="Right" > <font face="Verdana" size="1"><b><%=FormatNumber(SumCr,2)%></b></font></td>
		<td align="Right"  colspan="3" > <font face="Verdana" size="1"></font></td>
						
	</tr>
	</table>
	</div>
 </tbody>	

<!--#include file="bottom.Asp"-->
