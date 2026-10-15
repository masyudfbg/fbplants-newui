<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>FBPlants Reporting - HPP Kebun Harian</title>
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
}, false
);
</script>
<style type="text/css">
@media all {
#footer_note {width:65%;margin:auto;margin-top:20px;}
#footer_note .block { font-size:10px; float:left; width:20%; height:90px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center; font-weight:bold;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
}
@media print
{
#footer_note {margin-top:20px;}
#footer_note .block { font-size:10px; float:left; width:20%; height:90px; background-color:#fff; margin-left:1%; text-align:center;}
input[type="text"].isian {border:0px;box-shadow:none; padding:0px;font-size: 10px;color:#000;height: 20px;width:100%; text-align:center; font-weight:bold;}
hr{border-top:1px #000000 solid; border-bottom:0px; border-left:0px; border-right:0px;}
.printhide{display:none;}
@page { 
 size:Letter landscape;
}
}

</style>
</head>

<%
Dim Columnx : Columnx = 0
Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

Dim hppfrom,hppto,fcba,r1
dim sqldate,rsdate,startdate,sqlsbi,rssbi,startdatesbi,periode,t1,filterprd
dim rshpp,qq,filterstatus
cmdstyle="border-right: 1 solid #000000; border-bottom: 1 solid #000000"
cmdstyleTop="border-right: 1 solid #000000; border-bottom: 1 solid #000000; border-top: 1 solid #000000"
cmdstyleLeft="border-right: 1 solid #000000; border-bottom: 1 solid #000000; border-Left: 1 solid #000000"


	hppfrom 	= Request("hppfrom")
    hppto 		= Request("hppto")
    fcba 		= Request("fcba")
    r1   	    = Request("r1")
	
	q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & fcba & "' "
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

	

%>

<body>
	<div id="wrapper">
        <!--#include virtual="/print_excel_blob.asp"-->
		<table border="0" width="100%" id="product-table" class="xthf-site">
			<thead>
				<tr>
					
					<th colspan="15" style=" background-color:#FFFFFF; color:#000000; border:none">
						<div id="header-wrapper">
							<div id="head_judul" style="float:left; text-align:left;">
								<h4>PT. <%=CompanyName%></h4>
								<h4>Estate: <%=fcba%></h4>
															
							</div>
							<div id="subheader" class="centered" style="width:50%; margin:auto;">
								<h2>LAPORAN HPP KEBUN HARIAN</h2>
							</div>
						</div>
					</th>
				</tr>
						
				
				<tr>
					<th align="left" colspan="15" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">Periode: <%=(hppfrom) & "-" & hppto%></th>
				</tr>
						


	<%	
		qq = " WITH shipment AS ( " & _
			" SELECT " & _
			"   b.fccompanycode, " & _
        	"	ssh.fcba," & _
			"	ssd.postingdate tglkirim, " & _
			"	SUM(ssd.quantity) quantitysite, " & _
			"	SUM(ssd.acceptquantity) quantitypks " & _
			"		FROM salesshipmentdtl ssd " & _
			"		INNER JOIN salesshipmenthdr ssh " & _
			"			ON (ssd.salesshipmentno = ssh.salesshipmentno AND ssd.fcba = ssh.fcba) " & _
			"		INNER JOIN transporter tr  " & _
			"			ON (ssd.transporter = tr.fccode AND ssd.fcba = tr.fcba) " & _
			"		 INNER JOIN BUSINESSUNIT b ON (b.FCCODE=ssh.fcba) " & _
			"		WHERE ssh.fcba = '" & fcba & "' " & _
			"		AND ssd.POSTINGDATE BETWEEN TO_DATE('" & hppfrom & "','mm/dd/yyyy')  " & _
			"								AND TO_DATE('" & hppto & "','mm/dd/yyyy') " & _
			"		GROUP BY b.fccompanycode, ssh.fcba,ssd.postingdate " & _
			"	), " & _
			" lhp AS ( " & _
			"	SELECT b.fccompanycode, " & _
			"		la.unit AS fcba, " & _
			"		lh.docketdate AS tglkirim, " & _
			"		SUM(ld.ffbweight) AS quantitysite, " & _
			"		SUM(ld.ffbweight) AS quantitypks " & _
			"	FROM lhptikethdr lh " & _
			"	INNER JOIN lhpffbwbtrn ld " & _
			"		ON lh.docketid = ld.docketid " & _
			"	AND lh.fcba = ld.fcba " & _
			"	INNER JOIN lhpAreamst la " & _
			"		ON la.areaid = ld.areaid " & _
			"	AND la.fcba = ld.fcba " & _
			"	INNER JOIN businessunit b " & _
			"		ON b.fccode = la.unit " & _
			"	WHERE la.unit = '" & fcba & "' " & _
			"	AND lh.docketdate BETWEEN TO_DATE('" & hppfrom & "','mm/dd/yyyy') " & _
			"							AND TO_DATE('" & hppto & "','mm/dd/yyyy') " & _
			"	AND lh.productid = 2 " & _
			"	GROUP BY b.fccompanycode, la.unit, lh.docketdate " & _
			" ), " & _
			" shipment_final AS ( " & _
			"	SELECT * FROM shipment " & _
			"	UNION ALL " & _
			"	SELECT * FROM lhp " & _
			" ),"  & _
			"	biaya AS ( " & _
			"		SELECT  " & _
			"			b.fccompanycode, f.fcba,f.journaldate tglkirim, " & _
			"			SUM(CASE WHEN j.JOB_FORBUDGETCODE = '101010' THEN f.debet ELSE 0 END) perawatantm, " & _
			"			SUM(CASE WHEN j.JOB_FORBUDGETCODE = '101020' THEN f.debet ELSE 0 END) pemupukantm, " & _
			"			SUM(CASE WHEN j.JOB_FORBUDGETCODE = '101030' THEN f.debet ELSE 0 END) panendantransport, " & _
			"			SUM(CASE WHEN j.JOB_FORBUDGETCODE = '101040' THEN f.debet ELSE 0 END) sarpastm, " & _
			"			SUM(CASE WHEN j.JOB_FORBUDGETCODE = '101050' THEN f.debet ELSE 0 END) umum " & _
			"		FROM GLEDGERDTL f " & _
			"		INNER JOIN jobmaster j " & _
			"			ON f.coa = j.fccode " & _
			"		INNER JOIN BUSINESSUNIT b ON (b.FCCODE=f.fcba) " & _
			"		WHERE journaldate BETWEEN TO_DATE('" & hppfrom & "','mm/dd/yyyy')  " & _
			"								AND TO_DATE('" & hppto & "','mm/dd/yyyy') " & _
			"		AND f.fcba = '" & fcba & "' " & _
			"		AND f.SOURCECODE IN('OPR-BKM','MAM-SIS','FAC-VRA','OPR-WIP','FAC-BNK') " & _
			"		GROUP BY b.fccompanycode, f.fcba,f.journaldate " & _
			"	) " & _
			"	SELECT  " & _
			"       b.fccompanycode," & _
    		"		b.fcba," & _
			"		NVL(s.tglkirim, b.tglkirim) tglkirim,"  & _
			"		NVL(s.quantitysite,0) quantitysite, " & _
			"		NVL(s.quantitypks,0) quantitypks, " & _
			"		NVL(s.quantitysite,0) - NVL(s.quantitypks,0) AS varian," & _
			"		NVL(b.perawatantm,0) perawatantm, " & _
			"		NVL(b.pemupukantm,0) pemupukantm, " & _
			"		NVL(b.panendantransport,0) panendantransport, " & _
			"		NVL(b.sarpastm,0) sarpastm, " & _
			"		NVL(b.umum,0) umum, " & _
			"		(NVL(b.perawatantm,0) " & _
			"		+ NVL(b.pemupukantm,0) " & _
			"		+ NVL(b.panendantransport,0) " & _
			"		+ NVL(b.sarpastm,0) " & _
			"		+ NVL(b.umum,0)) total " & _
			"	FROM shipment_final s " & _
			"	FULL OUTER JOIN biaya b " & _
			"	ON s.tglkirim = b.tglkirim " & _
			"	ORDER BY NVL(s.tglkirim, b.tglkirim) ASC NULLS FIRST"
		''	response.write qq
			set rshpp = connStr.Execute(qq)
	%>
	<tr>
			<th rowspan="2">No</th>
		    <th rowspan="2">Company</th>
		    <th rowspan="2">FCBA</th>
		    <th rowspan="2">Tanggal Kirim</th>
			<th colspan="3">TBS (KG)</th>

			<th rowspan="2">Perawatan TM</th>
			<th rowspan="2">Pemupukan TM</th>
			<th rowspan="2">Panen dan Transport</th>
			<th rowspan="2">Sarpas TM</th>
			<th rowspan="2">Umum</th>
			<th rowspan="2">Total</th>
			<th rowspan="2">HPP / Kg Kebun</th>
			<th rowspan="2">HPP / Kg PKS</th>
			
	</tr>
	<tr>
		<th>Timbangan Kebun</th>
		<th>Timbangan PKS</th>
		<th>Selisih</th>
	</tr>
	</thead>
	<%
		dim rpjtkebun,rpjtpks,totaljtkebun,totaljtpks,totalvarian,totalperawatantm,totalpemupukantm,totalpanentransporttm,totalsarpastm,totalumum,totalsemua,totalrpjtkebun,totalrpjtpks
		if Not rshpp.eof then
		no=1
		totaljtkebun = 0
		totaljtpks   = 0
		totalvarian  = 0 
		totalperawatantm = 0 
		totalpemupukantm = 0
		totalpanentransporttm = 0
		totalsarpastm = 0 
		totalumum = 0
		totalsemua = 0 
		totalrpjtkebun = 0
		totalrpjtpks = 0
		While Not rshpp.EOF
	''	rpjtkebun = cdbl(rshpp("total")) / cdbl(rshpp("quantitysite"))
	''	rpjtpks   = cdbl(rshpp("total")) / cdbl(rshpp("quantitypks"))
		If cdbl(rshpp("quantitysite")) <> 0 Then
			rpjtkebun = cdbl(rshpp("total")) / cdbl(rshpp("quantitysite"))
		Else
			rpjtkebun = 0
		End If

		If cdbl(rshpp("quantitypks")) <> 0 Then
			rpjtpks = cdbl(rshpp("total")) / cdbl(rshpp("quantitypks"))
		Else
			rpjtpks = 0
		End If
	%>

	<tr>
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=no%></td>					
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rshpp("fccompanycode")%></td>			
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rshpp("fcba")%></td>				
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rshpp("tglkirim")%></td>				
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rshpp("quantitysite"),2)%></td>				
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rshpp("quantitypks"),2)%></td>	
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rshpp("varian"),2)%></td>		
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rshpp("perawatantm"),2)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rshpp("pemupukantm"),2)%></td>			
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rshpp("panendantransport"),2)%></td>					
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rshpp("sarpastm"),2)%></td>
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rshpp("umum"),2)%></td>			
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rshpp("total"),2)%></td>			
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rpjtkebun,2)%></td>			
			<td align="right" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rpjtpks,2)%></td>			
			
	</tr>



	<%
		totaljtkebun = cdbl(totaljtkebun) + cdbl(rshpp("quantitysite"))
		totaljtpks   = cdbl(totaljtpks) + cdbl(rshpp("quantitypks"))
		totalvarian  = cdbl(totalvarian) + cdbl(rshpp("varian"))
		totalperawatantm = cdbl(totalperawatantm) + cdbl(rshpp("perawatantm"))
		totalpemupukantm = cdbl(totalpemupukantm) + cdbl(rshpp("pemupukantm"))
		totalpanentransporttm = cdbl(totalpanentransporttm) + cdbl(rshpp("panendantransport"))
		totalsarpastm = cdbl(totalsarpastm) + cdbl(rshpp("sarpastm"))
		totalumum = cdbl(totalumum) + cdbl(rshpp("umum"))
		totalsemua = cdbl(totalsemua) + cdbl(rshpp("total"))
		If cdbl(totaljtkebun) <> 0 Then
			totalrpjtkebun = cdbl(totalsemua) / cdbl(totaljtkebun)
		Else
			totalrpjtkebun = 0
		End If
		If cdbl(totaljtpks) <> 0 Then
			totalrpjtpks = cdbl(totalsemua) / cdbl(totaljtpks)
		Else
			totalrpjtpks = 0
		End If
		
		

		rshpp.MoveNext
		no=no+1
		Wend
		rshpp.close()
		rshpp=null
		END IF
	%>


<tbody>


		<tr bgcolor="#F4F4A0">
			<td align="Center" style="<%=cmdstyle%>" colspan= "4"><font face="Verdana" size="1"><strong>T o t a l</strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(totaljtkebun,2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(totaljtpks,2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(totalvarian,2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(totalperawatantm,2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(totalpemupukantm,2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(totalpanentransporttm,2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(totalsarpastm,2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(totalumum,2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(totalsemua,2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(totalrpjtkebun,2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(totalrpjtpks,2)%></strong></td>
		</tr>
	</tbody>
</table>
<!--#include virtual="/Bottom_Report.asp"---------->

		</DIV>
	</BODY>
</HTML>	