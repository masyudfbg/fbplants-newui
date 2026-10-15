<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Laporan Tanaman Rekap Harian</title>

<script language="JavaScript" src='/js/x.js'></script>
	<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
	<script language="JavaScript">
	xAddEventListener(window, 'load',
	  function() {
		new xTableHeaderFixed('xthf-site', window);
	  }, false
	);

	</script>


<style type="text/css">
@media print
{
.printhide{display:none;}
}

.tebal{
	font-weight:bold;
}

.ratakiri{
	text-align:left;
}
.ratakanan{
	text-align:right;
}
.ratatengah{
	text-align:center;
}
.lebarket{
	width:780px;
}

</style>
</head>

<%



Dim tbl,doc,nday,txtSQL,rsdata,TGLAWAL,TGLAKHIR
Dim TTLSEMUA,TTLARJ,TTLBRS,TTLBIA,TTLKAA,TTLGAS,TTLKSR,TTLBRP,TTLCAE,TTLESR



cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

TGLAWAL		= trim(REQUEST("tglawal"))
TGLAKHIR	= trim(REQUEST("tglakhir"))


		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
			' response.write q
			set rsCat = connStr.Execute(q)
			if not rsCat.eof then
			   fcbaname = rsCat("fcName")
			   CompanyName=rsCat("FCCOMPANYNAME")
			   MgrName = rsCat("FCMANAGER")
			   KTUName = rsCat("FCKTU")
			   CropName=rsCat("FCCROP")
			   FcbaType=rsCat("fcType")
			end if

	txtSQL = "SELECT G.ID,G.URAIAN GROUPMST,M.URAIAN ITEM,M.SATUAN, " & _
    " SUM(CASE WHEN T.FCBA='GWE' then T.JUMLAH else 0 end) AS GWE, " & _
    " SUM(CASE WHEN T.FCBA='PSE' then T.JUMLAH else 0 end) AS PSE , " & _
    " SUM(CASE WHEN T.FCBA='SKE' then T.JUMLAH else 0 end) AS SKE , " & _
    " SUM(CASE WHEN T.FCBA='ARM' then T.JUMLAH else 0 end) AS ARM, " & _
    " SUM(CASE WHEN T.FCBA='BRE' then T.JUMLAH else 0 end) AS BRE, " & _
    " SUM(CASE WHEN T.FCBA='CDE' then T.JUMLAH else 0 end) AS CDE, " & _
    " SUM(CASE WHEN T.FCBA='GHE' then T.JUMLAH else 0 end) AS GHE, " & _
    " SUM(CASE WHEN T.FCBA='EBE' then T.JUMLAH else 0 end) AS EBE, " & _
	" SUM(CASE WHEN T.FCBA='TPE' then T.JUMLAH else 0 end) AS TPE, " & _
	" SUM(CASE WHEN T.FCBA='SFE' then T.JUMLAH else 0 end) AS SFE, " & _
	" SUM(CASE WHEN T.FCBA='RBE' then T.JUMLAH else 0 end) AS RBE, " & _
	" SUM(CASE WHEN T.FCBA='DLE' then T.JUMLAH else 0 end) AS DLE, " & _
	" SUM(CASE WHEN T.FCBA='MREI' then T.JUMLAH else 0 end) AS MREI, " & _
	" SUM(CASE WHEN T.FCBA='RMIE' then T.JUMLAH else 0 end) AS RMIE, " & _
	" SUM(CASE WHEN T.FCBA='KSR' then T.JUMLAH else 0 end) AS KSR, " & _
	" SUM(CASE WHEN T.FCBA='BRP' then T.JUMLAH else 0 end) AS BRP, " & _
	" SUM(CASE WHEN T.FCBA='CAE' then T.JUMLAH else 0 end) AS CAE, " & _
	" SUM(CASE WHEN T.FCBA='ESR' then T.JUMLAH else 0 end) AS ESR " & _
	" from TANAMANTRN T " & _
	" INNER JOIN TANAMANMST M ON (T.KODE=M.KODE) " & _
	" INNER JOIN TANAMANGROUP G ON (G.ID=M.GROUPMST) " & _
	" WHERE T.TANGGAL BETWEEN TO_DATE('" & TGLAWAL & "','mm/dd/yyyy') AND TO_DATE('" & TGLAKHIR & "','mm/dd/yyyy')  " & _
	" GROUP BY G.ID,G.URAIAN, T.KODE, M.URAIAN,M.SATUAN " & _
	" ORDER BY G.ID"
 	Set rsdata = Server.CreateObject("ADODB.Recordset")
	rsdata.CursorType = 3
	rsdata.CursorLocation = 3
	rsdata.Open txtSQL, connstr

%>

<body>
	<div id="wrapper">	
		<div id="header-wrapper">
		<div class="printhide">
        <!--#include file="print_excel.asp"-->
		<table border="1" width="100%" id="product-table" class="xthf-site">
			<thead>
				<tr>
					
					<th colspan="35" style=" background-color:#FFFFFF; color:#000000; border:none">
						<div id="header-wrapper">
							<div id="head_judul" style="float:left; text-align:left;">
								<h4>FBG Group</h4>

								<h5>SEMUA PT.</h5>
							</div>
							
						</div>
						<div id="subheader" class="ratakiri">
								
							<tr>
								<th align="left" colspan="3" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:left;">LAPORAN HARIAN (REKAP)</th><th colspan="34" style=" background-color:#FFFFFF; color:#000000; border:none;" class="ratakanan">PERIODE: <%=tahunbulan%></th>
							</tr>
							<tr>	
								
							</tr>
						</div>
					</th>
				</tr>
				
	
	<tr>
		 <th style="color:#fff;" class=lebarket>ITEM</th>
		 <th style="color:#fff;" class=lebarket>URAIAN</th>
		 <th style="color:#fff;">UOM</th> 
		 <th style="color:#fff;">ALL</th> 
		 <th style="color:#fff;">ARJ</th> 
		 <th style="color:#fff;">BRS</th> 
		 <th style="color:#fff;">BIA</th> 
		 <th style="color:#fff;">KAA</th> 
		 <th style="color:#fff;">GAS</th> 
		 <th style="color:#fff;">KSR</th> 
		 <th style="color:#fff;">BRP</th> 
		 <th style="color:#fff;">CAE</th> 
		 <th style="color:#fff;">ESR</th> 
	</tr>
	
	
</thead>
	<tbody>
	<%
	
		while not rsdata.eof

		
		TTLSEMUA = cdbl(rsdata("GWE"))+cdbl(rsdata("PSE"))+cdbl(rsdata("SKE"))+cdbl(rsdata("BRE"))+cdbl(rsdata("CDE"))+cdbl(rsdata("GHE"))+cdbl(rsdata("EBE"))+cdbl(rsdata("TPE"))+cdbl(rsdata("SFE"))+cdbl(rsdata("RBE"))+ cdbl(rsdata("DLE"))+cdbl(rsdata("MREI"))+cdbl(rsdata("RMIE"))+cdbl(rsdata("KSR"))+cdbl(rsdata("BRP"))+cdbl(rsdata("CAE"))+cdbl(rsdata("ESR"))
		TTLARJ	 = cdbl(rsdata("GWE"))+cdbl(rsdata("PSE"))+cdbl(rsdata("SKE"))
		TTLBRS	 = cdbl(rsdata("BRE"))
		TTLBIA	 = cdbl(rsdata("CDE"))+cdbl(rsdata("GHE"))+cdbl(rsdata("EBE"))
		TTLKAA 	 = cdbl(rsdata("TPE"))+cdbl(rsdata("SFE"))+cdbl(rsdata("RBE"))+cdbl(rsdata("DLE"))
		TTLGAS	 = cdbl(rsdata("MREI"))+cdbl(rsdata("RMIE"))
		TTLKSR	 = cdbl(rsdata("KSR"))
		TTLBRP	 = cdbl(rsdata("BRP"))
		TTLCAE	 = cdbl(rsdata("CAE"))
		TTLESR   = cdbl(rsdata("ESR"))

		
	%>
	<tr class="<%=cClass%>">
		<td align="left" Rowspan="" class="lebarket"><%=rsdata("GROUPMST")%></td> 
		<td align="left"><%=rsdata("ITEM")%></td>
		<td align="center"><%=rsdata("SATUAN")%></td>
		<td align="right"><%=FormatNumber(TTLSEMUA,2)%></td>
		<td align="right"><%=FormatNumber(TTLARJ,2)%></td>
		<td align="right"><%=FormatNumber(TTLBRS,2)%></td>
		<td align="right"><%=FormatNumber(TTLBIA,2)%></td>
		<td align="right"><%=FormatNumber(TTLKAA,2)%></td>
		<td align="right"><%=FormatNumber(TTLGAS,2)%></td>
		<td align="right"><%=FormatNumber(TTLKSR,2)%></td>
		<td align="right"><%=FormatNumber(TTLBRP,2)%></td>
		<td align="right"><%=FormatNumber(TTLCAE,2)%></td>
		<td align="right"><%=FormatNumber(TTLESR,2)%></td>
		
	</tr>
	
	
	<%
	
	
		
		NO = NO + 1	
		rsdata.MoveNext
		WEND
		rsdata.Close
		set rsdata = Nothing
	
	%>

	</tbody>
	</div>
</table>
<!--#include file="Bottom_Report.asp"---------->

		</DIV>
	</div>
</HTML>	