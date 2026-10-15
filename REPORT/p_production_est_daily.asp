<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%	
	Dim sqlprod, rsprod,fcba,q,NO,divisi,n
	Dim fcbaname,CompanyName,MgrName,KTUName,CropName
	Dim tglawal,tglakhir
	Dim bln,tgl,tahun,blntahun,filterdiv,bulan,prd
	fcba 		= trim(request("businessunit"))
	divisi 		= trim(request("divisi"))
	prd 		= trim(request("prd"))
	bulan 		= trim(request("bulan"))
	tahun 		= trim(request("tahun"))
	

	blntahun	= bulan&tahun
	if prd = "P" then	
		q = "Select start_date,end_date from empalldedstatus where fcba='" & fcba & "' and year=" & tahun & " and month=" & bulan
		set rsCat = connStr.Execute(q)
		if not rsCat.eof then
			tglawal  = rsCat("start_date")
			tglakhir = rsCat("end_date")
		end if
	else
		q = "Select to_date('" & bulan & "/01/" & tahun & "','MM/DD/YYYY') start_date, LAST_DAY(to_date('" & bulan & "/01/" & tahun & "','MM/DD/YYYY')) end_date from dual "
		
		set rsCat = connStr.Execute(q)
		if not rsCat.eof then
			tglawal  = rsCat("start_date")
			tglakhir = rsCat("end_date")
		end if
	end if


	
	rsCat.Close
	set rsCat = Nothing
			
	
	q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & fcba & "' "
		' response.write q
		set rsCat = connStr.Execute(q)
		if not rsCat.eof then
			fcbaname = rsCat("fcName")
			CompanyName=rsCat("FCCOMPANYNAME")
			MgrName = rsCat("FCMANAGER")
			KTUName = rsCat("FCKTU")
			CropName=rsCat("FCCROP")
		end if
	rsCat.Close
	set rsCat = Nothing
	
	
	sqlprod = " BEGIN " & _
		" bkmproduksi_api.prosess(to_date('" & tglawal & "','MM/DD/YYYY'),to_date('" & tglakhir & "','MM/DD/YYYY'),'" & fcba & "');" & _
		" END;"
	ExecuteAndLog "p_production_est_daily.asp", sqlprod
	

	if trim(divisi) <> "all" then
		filterdiv = " WHERE t.DIVISION = '" & divisi & "'"
	end if

	sqlprod = "SELECT T.DIVISION,T.FIELDCODE,NVL(T.LUAS_GROSS,0)LUAS_GROSS,NVL(T.PRODUCTIVESTAND,0)PRODUCTIVESTAND,NVL(T.HECTARAGEPLANTED ,0)HECTARAGEPLANTED,NVL(SUM(T.HAHI),0)HAHI,NVL(SUM(T.HASHI),0)HASHI,NVL(SUM(T.HASBI),0)HASBI, " & _
	" NVL(SUM(T.HKHI),0)HKHI,NVL(SUM(T.HKSHI),0)HKSHI,NVL(SUM(T.HKSBI),0)HKSBI, " & _
	" NVL(SUM(T.JJGPANENHI),0)JJGPANENHI,NVL(SUM(T.JJGPANENSHI),0)JJGPANENSHI,NVL(SUM(T.JJGPANENSBI),0)JJGPANENSBI, " & _
	" DECODE(NVL(SUM(T.JJGPANENHI),0),0,0,(SUM(T.TONASEESTHI))/NVL(SUM(T.JJGPANENHI),0))BJRHI,DECODE(NVL(SUM(T.JJGPANENSHI),0),0,0,(SUM(T.TONASEESTSHI))/NVL(SUM(T.JJGPANENSHI),0))BJRSHI,DECODE(NVL(SUM(T.JJGPANENSBI),0),0,0,(SUM(T.TONASEESTSBI))/NVL(SUM(T.JJGPANENSBI),0))BJRSBI, " & _
	" NVL(SUM(T.TONASEESTHI),0)TONASEESTHI,NVL(SUM(T.TONASEESTSHI),0)TONASEESTSHI,NVL(SUM(T.TONASEESTSBI),0)TONASEESTSBI, " & _
	" NVL(SUM(T.JJGPKSHI),0)JJGPKSHI,NVL(SUM(T.JJGPKSSHI),0)JJGPKSSHI,NVL(SUM(T.JJGPKSSBI),0)JJGPKSSBI, " & _
	" NVL(SUM(T.TONASEPKSHI),0)TONASEPKSHI,NVL(SUM(T.TONASEPKSSHI),0)TONASEPKSSHI,NVL(SUM(T.TONASEPKSSBI),0)TONASEPKSSBI, " & _
	" DECODE(NVL(SUM(T.HECTARAGEPLANTED),0),0,0,(SUM(T.HAHI))/NVL(T.HECTARAGEPLANTED,0)) ROTASIHI,DECODE(NVL(SUM(T.HECTARAGEPLANTED),0),0,0,(SUM(T.HASHI))/NVL(T.HECTARAGEPLANTED,0)) ROTASISHI,DECODE(NVL(SUM(T.HECTARAGEPLANTED),0),0,0,SUM(T.HASBI)/NVL(T.HECTARAGEPLANTED,0)) ROTASISBI, " & _
	" (NVL(SUM(T.JJGPANENHI),0)-NVL(SUM(T.JJGPKSHI),0))RESTANJJGHI,(NVL(SUM(T.JJGPANENSHI),0)-NVL(SUM(T.JJGPKSSHI),0))RESTANJJGSHI,(NVL(SUM(T.JJGPANENSBI),0)-NVL(SUM(T.JJGPKSSBI),0))RESTANJJGSBI " & _
	" FROM( " & _
	"	SELECT P.DIVISION,P.FIELDCODE,F.LUAS_GROSS,F.PRODUCTIVESTAND,F.HECTARAGEPLANTED,NVL(SUM(P.OUTPUT_HI),0)HAHI,0 HASHI,0 HASBI, " & _
	"	NVL(SUM(P.MANDAYS_HI),0)HKHI,0 HKSHI,0 HKSBI, " & _
	"	NVL(SUM(P.JJGHI),0)JJGPANENHI,0 JJGPANENSHI,0 JJGPANENSBI, " & _
	"	0 BJRHI,0 BJRSHI,0 BJRSBI, " & _
	"	NVL(SUM(P.TONASE_HI),0)TONASEESTHI,0 TONASEESTSHI,0 TONASEESTSBI, " & _
	"	NVL(SUM(P.POM_HI),0)JJGPKSHI,0 JJGPKSSHI,0 JJGPKSSBI, " & _
	"	NVL(SUM(P.POM_TONASE_HI),0)TONASEPKSHI,0 TONASEPKSSHI,0 TONASEPKSSBI " & _
	"	FROM PRODUCTION_EST P " & _
	"	LEFT JOIN FIELD_HIST_TAB F ON P.DIVISION = F.DIVISION AND P.FIELDCODE=F.FCCODE AND P.FCBA=F.FCBA " & _
	"	WHERE P.FCBA='" & fcba & "' AND P.FDDATE=TO_DATE('" & tglakhir & "','MM/DD/YYYY') " & _
	"	GROUP BY P.DIVISION,P.FIELDCODE,F.LUAS_GROSS,F.PRODUCTIVESTAND,F.HECTARAGEPLANTED " & _
	"	UNION ALL " & _
	"	SELECT P.DIVISION,P.FIELDCODE,F.LUAS_GROSS,F.PRODUCTIVESTAND,F.HECTARAGEPLANTED,0 HAHI,NVL(SUM(P.OUTPUT_HI),0) HASHI,0 HASBI, " & _
	"	0 HKHI,NVL(SUM(P.MANDAYS_HI),0)HKSHI,0 HKSBI, " & _
	"	0 JJGPANENHI,NVL(SUM(P.JJGHI),0)JJGPANENSHI,0 JJGPANENSBI, " & _
	"	0 BJRHI,0 BJRSHI,0 BJRSBI, " & _
	"	0 TONASEESTHI,NVL(SUM(P.TONASE_HI),0)TONASEESTSHI,0 TONASEESTSBI, " & _
	"	0 JJGPKSHI,NVL(SUM(P.POM_HI),0)JJGPKSSHI,0 JJGPKSSBI, " & _
	"	0 TONASEPKSHI,NVL(SUM(P.POM_TONASE_HI),0)TONASEPKSSHI,0 TONASEPKSSBI " & _
	"	FROM PRODUCTION_EST P " & _
	"	LEFT JOIN FIELD_HIST_TAB F ON P.DIVISION = F.DIVISION AND P.FIELDCODE=F.FCCODE AND P.FCBA=F.FCBA " & _
	"	WHERE P.FCBA='" & fcba & "' AND P.FDDATE BETWEEN TO_DATE('" & tglawal & "','MM/DD/YYYY') AND TO_DATE('" & tglakhir & "','MM/DD/YYYY') " & _
	"	GROUP BY P.DIVISION,P.FIELDCODE,F.LUAS_GROSS,F.PRODUCTIVESTAND,F.HECTARAGEPLANTED " & _
	"	UNION ALL " & _
	"	SELECT P.DIVISION,P.FIELDCODE,F.LUAS_GROSS,F.PRODUCTIVESTAND,F.HECTARAGEPLANTED,0 HAHI,0 HASHI,NVL(SUM(P.OUTPUT_HI),0) HASBI, " & _
	"	0 HKHI,0 HKSHI,NVL(SUM(P.MANDAYS_HI),0)HKSBI, " & _
	"	0 JJGPANENHI,0 JJGPANENSHI,NVL(SUM(P.JJGHI),0) JJGPANENSBI, " & _
	"	0 BJRHI,0 BJRSHI,0 BJRSBI, " & _
	"	0,0 TONASEESTSHI,NVL(SUM(P.TONASE_HI),0)TONASEESTSBI, " & _
	"	0 JJGPKSHI,0 JJGPKSSHI,NVL(SUM(P.POM_HI),0) JJGPKSSBI, " & _
	"	0 TONASEPKSHI,0 TONASEPKSSHI,NVL(SUM(P.POM_TONASE_HI),0) TONASEPKSSBI " & _
	"	FROM PRODUCTION_EST P " & _
	"	INNER JOIN FIELD_HIST_TAB F ON P.DIVISION = F.DIVISION AND P.FIELDCODE=F.FCCODE AND P.FCBA=F.FCBA " & _
	"	WHERE P.FCBA='" & fcba & "' AND P.FDDATE BETWEEN TO_DATE('01/01/" & tahun & "','MM/DD/YYYY') AND TO_DATE('" & tglakhir & "','MM/DD/YYYY')  " & _
	"	GROUP BY P.DIVISION,P.FIELDCODE,F.LUAS_GROSS,F.PRODUCTIVESTAND,F.HECTARAGEPLANTED " & _
	" )T " & _
	filterdiv & _
	" GROUP BY T.DIVISION,T.FIELDCODE,T.LUAS_GROSS,T.PRODUCTIVESTAND,T.HECTARAGEPLANTED " & _
	" ORDER BY T.DIVISION,T.FIELDCODE"
	
	'Response.write sqlprod
	Set rsprod = Server.CreateObject("ADODB.Recordset")
	rsprod.CursorType = 3
	rsprod.CursorLocation = 3
	rsprod.Open sqlprod, connstr
	


%>

<html>

<head>
<title>ePlaS Reporting - Laporan Harian Produksi</title>
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

</head>

<body>
	<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
<form name="frmEntry" method="post" action="p_ProductionDaily.asp">

<!--
	<table id="menu-table" class="centered-margin">
		<tr>
			<td align="center"><a href="/login.asp">Back to Home</a></td>
			<td align="center"><a href="/report/Dashboard.asp">Back to Report</a></td>
		</tr>
	</table>
	-->

<table border="0" width="100%" id="product-table" class="xthf-site">
<thead>
<tr>
<th colspan="38" style="background:#ffffff; border-top:1px #ffffff solid; border-right:1px #ffffff solid; border-left:1px #ffffff solid;">
<div id="header-wrapper">
		<div id="header" style="text-align:left">
			<% if  CompanyName <>  "" then %>
				<h4>PT.<%=CompanyName%></h4> 
			<% End if%>
			<% If  FcbaName <> ""  then %>
				   <h5><%=FcbaName%></h5>   
			<% End if%>
		</div>
		<div id="subheader"class="centered">
			<h2>LAPORAN HARIAN PRODUKSI</h2>
		
				<h5>SAMPAI DENGAN : <%=tglakhir%>, KEBUN : <%=fcba%></h5>
		
			<h3></br></h3>
			<h3></br></h3>
		</div>
	</div>


</th>
</tr>
	<tr>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">NO</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">DIVISI</th>
		<th rowspan="2" style="background-color:#00A300;color:#fff;">BLOK</th>
		<th colspan="2" style="background-color:#00A300;color:#fff;">PRODUKSI</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">HA</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">HK</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">JJG (Panen)</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">BJR</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">TONASE (KEBUN)</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">JJG (PKS)</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">TONASE (PKS)</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">ROTASI</th>
		<th colspan="3" style="background-color:#00A300;color:#fff;">RESTAN</th>
	</tr>
	
	
	<tr>

	<!-- Produksi-->
		<th  style="background-color:#00A300;color:#fff;" >Luas</th>
		<th  style="background-color:#00A300;color:#fff;" >Pokok</th>

	<!-- HA-->	
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
	<!-- HK-->	
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
	<!-- JJG Panen-->	
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
	<!-- BJR-->		
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
	<!-- TONASE KEBUN-->	
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
	<!-- JJG PKS-->	
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
	<!-- TON PKS-->	
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>
  
	<!-- ROTASI-->		
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>

	<!-- RESTAN-->
		<th  style="background-color:#00A300;color:#fff;" >HI</th>
		<th  style="background-color:#00A300;color:#fff;" >SHI</th>
		<th  style="background-color:#00A300;color:#fff;" >SBI</th>

		
	</tr>
        </thead>
 <tbody>
<%

	dim nfieldha,nfieldpokok
	dim nhaHI,nhaSHI,nhaSBI
	dim nhkHI,nhkSHI,nhkSBI
	dim njjgHI,njjgSHI,njjgSBI
	dim nbjrHI,nbjrSHI,nbjrSBI
	dim ntonaseHI,ntonaseSHI,ntonaseSBI
	dim njjgpksHI,njjgpksSHI,njjgpksSBI
	dim ntonasepksHI,ntonasepksSHI,ntonasepksSBI
	dim nrotasiHI,nrotasiSHI,nrotasiSBI
	dim nrestanHI,nrestanSHI,nrestanSBI
	if not rsprod.EOF then
		NO = 1
		nfieldha = 0
		nfieldpokok = 0
		nhaHI = 0
		nhaSHI = 0
		nhaSBI = 0
		nhkHI = 0
		nhkSHI = 0
		nhkSBI = 0
		njjgHI = 0
		njjgSHI = 0
		njjgSBI = 0
		nbjrHI = 0
		nbjrSHI = 0
		nbjrSBI = 0
		ntonaseHI = 0
		ntonaseSHI = 0
		ntonaseSBI = 0
		njjgpksHI = 0
		njjgpksSHI = 0
		njjgpksSBI = 0
		ntonasepksHI = 0
		ntonasepksSHI = 0
		ntonasepksSBI = 0
		nrotasiHI = 0
		nrotasiSHI = 0
		nrotasiSBI = 0
		nrestanHI = 0
		nrestanSHI = 0
		nrestanSBI = 0
		while not rsprod.EOF
			
%>
			<tr class="">
				<!-- DIVISI -->
				<td align="right"><%=NO%></td>
				<td align="right"><%=FbpNum2((rsprod("DIVISION")))%></td>
				<td align="right"><%=(rsprod("FIELDCODE"))%></td>

				<!-- PRODUKSI -->
				<td align="right"><%=FbpNum2((rsprod("LUAS_GROSS")))%></td>
				<td align="right"><%=FbpNum2((rsprod("PRODUCTIVESTAND")))%></td>
				
				<!-- HA -->
				<td align="right"><%=FormatNumber(rsprod("HAHI"),2)%></td>
				<td align="right"><%=FormatNumber(rsprod("HASHI"),2)%></td>
				<td align="right"><%=FormatNumber(rsprod("HASBI"),2)%></td>

				<!-- HK -->
				<td align="right"><%=FormatNumber(rsprod("HKHI"),2)%></td>
				<td align="right"><%=FormatNumber(rsprod("HKSHI"),2)%></td>
				<td align="right"><%=FormatNumber(rsprod("HKSBI"),2)%></td>


				<!-- JJG -->
				<td align="right"><%=FormatNumber(rsprod("JJGPANENHI"),2)%></td>
				<td align="right"><%=FormatNumber(rsprod("JJGPANENSHI"),2)%></td>
				<td align="right"><%=FormatNumber(rsprod("JJGPANENSBI"),2)%></td>
						
				<!-- BJR -->
				<td align="right"><%=FormatNumber(rsprod("BJRHI"),2)%></td>
				<td align="right"><%=FormatNumber(rsprod("BJRSHI"),2)%></td>
				<td align="right"><%=FormatNumber(rsprod("BJRSBI"),2)%></td>

				<!-- Tonase KEBUN -->
				<td align="right"><%=FormatNumber(cdbl(rsprod("TONASEESTHI")),2)%></td>
				<td align="right"><%=FormatNumber(cdbl(rsprod("TONASEESTSHI")),2)%></td>
				<td align="right"><%=FormatNumber(cdbl(rsprod("TONASEESTSBI")),2)%></td>

				<!-- JJG PKS-->
				<td align="right"><%=FormatNumber(rsprod("JJGPKSHI"),2)%></td>
				<td align="right"><%=FormatNumber(rsprod("JJGPKSSHI"),2)%></td>
				<td align="right"><%=FormatNumber(rsprod("JJGPKSSBI"),2)%></td>
				
				<!-- TONASE PKS -->
				
				<td align="right"><%=FormatNumber(rsprod("TONASEPKSHI"),2)%></td>
				<td align="right"><%=FormatNumber(rsprod("TONASEPKSSHI"),2)%></td>
				<td align="right"><%=FormatNumber(rsprod("TONASEPKSSBI"),2)%></td>
			
				<!-- ROTASI -->
				<td align="right"><%=FormatNumber(rsprod("ROTASIHI"),2)%></td>
				<td align="right"><%=FormatNumber(rsprod("ROTASISHI"),2)%></td>
				<td align="right"><%=FormatNumber(rsprod("ROTASISBI"),2)%></td>
				
				<!-- RESTAN -->
				<td align="right"><%=FormatNumber(rsprod("RESTANJJGHI"),2)%></td>
				<td align="right"><%=FormatNumber(rsprod("RESTANJJGSHI"),2)%></td>
				<td align="right"><%=FormatNumber(rsprod("RESTANJJGSBI"),2)%></td>
								
				
			</tr>
<%
			NO = NO + 1
			nfieldha = cdbl(nfieldha) + cdbl(rsprod("LUAS_GROSS"))
			nfieldpokok = cdbl(nfieldpokok) + cdbl(rsprod("PRODUCTIVESTAND"))
			nhaHI = cdbl(nhaHI) + cdbl(rsprod("HAHI"))
			nhaSHI = cdbl(nhaSHI) + cdbl(rsprod("HASHI"))
			nhaSBI = cdbl(nhaSBI) + cdbl(rsprod("HASBI"))
			nhkHI = cdbl(nhkHI) + cdbl(rsprod("HKHI"))
			nhkSHI = cdbl(nhkSHI) + cdbl(rsprod("HKSHI"))
			nhkSBI = cdbl(nhkSBI) + cdbl(rsprod("HKSBI"))
			njjgHI = cdbl(njjgHI) + cdbl(rsprod("JJGPANENHI"))
			njjgSHI = cdbl(njjgSHI) + cdbl(rsprod("JJGPANENSHI"))
			njjgSBI = cdbl(njjgSBI) + cdbl(rsprod("JJGPANENSBI"))
			nbjrHI = cdbl(nbjrHI) + cdbl(rsprod("BJRHI"))
			nbjrSHI = cdbl(nbjrSHI) + cdbl(rsprod("BJRSHI"))
			nbjrSBI = cdbl(nbjrSBI) + cdbl(rsprod("BJRSBI"))
			ntonaseHI = cdbl(ntonaseHI) + cdbl(rsprod("TONASEESTHI"))
			ntonaseSHI = cdbl(ntonaseSHI) + cdbl(rsprod("TONASEESTSHI"))
			ntonaseSBI = cdbl(ntonaseSBI) + cdbl(rsprod("TONASEESTSBI"))
			njjgpksHI = cdbl(njjgpksHI) + cdbl(rsprod("JJGPKSHI"))
			njjgpksSHI = cdbl(njjgpksSHI) + cdbl(rsprod("JJGPKSSHI"))
			njjgpksSBI = cdbl(njjgpksSBI) + cdbl(rsprod("JJGPKSSBI"))
			ntonasepksHI = cdbl(ntonasepksHI) + cdbl(rsprod("TONASEPKSHI"))
			ntonasepksSHI = cdbl(ntonasepksSHI) + cdbl(rsprod("TONASEPKSSHI"))
			ntonasepksSBI = cdbl(ntonasepksSBI) + cdbl(rsprod("TONASEPKSSBI"))
			nrotasiHI = cdbl(nrotasiHI) + cdbl(rsprod("ROTASIHI"))
			nrotasiSHI = cdbl(nrotasiSHI) + cdbl(rsprod("ROTASISHI"))
			nrotasiSBI = cdbl(nrotasiSBI) + cdbl(rsprod("ROTASISBI"))
			nrestanHI = cdbl(nrestanHI) + cdbl(rsprod("RESTANJJGHI"))
			nrestanSHI = cdbl(nrestanSHI) + cdbl(rsprod("RESTANJJGSHI"))
			nrestanSBI = cdbl(nrestanSBI) + cdbl(rsprod("RESTANJJGSBI"))
			rsprod.MoveNext			
		Wend
	end if
	rsprod.Close
	set rsprod = Nothing
	

%>
			<tr bgcolor="#F4F4A0" style="font-weight:bold">
			   
				<td align="center" colspan="3">TOTAL</td>
				<td align="right"><%=FormatNumber(nfieldha,2)%></td>
				<td align="right"><%=FormatNumber(nfieldpokok,2)%></td>
				<td align="right"><%=FormatNumber(nhaHI,2)%></td>
				<td align="right"><%=FormatNumber(nhaSHI,2)%></td>
				<td align="right"><%=FormatNumber(nhaSBI,2)%></td>
				<td align="right"><%=FormatNumber(nhkHI,2)%></td>
				<td align="right"><%=FormatNumber(nhkSHI,2)%></td>
				<td align="right"><%=FormatNumber(nhkSBI,2)%></td>
				<td align="right"><%=FormatNumber(njjgHI,2)%></td>
				<td align="right"><%=FormatNumber(njjgSHI,2)%></td>
				<td align="right"><%=FormatNumber(njjgSBI,2)%></td>
				<td align="right"><%=FormatNumber(nbjrHI,2)%></td>
				<td align="right"><%=FormatNumber(nbjrSHI,2)%></td>
				<td align="right"><%=FormatNumber(nbjrSBI,2)%></td>
				<td align="right"><%=FormatNumber(ntonaseHI,2)%></td>
				<td align="right"><%=FormatNumber(ntonaseSHI,2)%></td>
				<td align="right"><%=FormatNumber(ntonaseSBI,2)%></td>
				<td align="right"><%=FormatNumber(njjgpksHI,2)%></td>
				<td align="right"><%=FormatNumber(njjgpksSHI,2)%></td>
				<td align="right"><%=FormatNumber(njjgpksSBI,2)%></td>
				<td align="right"><%=FormatNumber(ntonasepksHI,2)%></td>
				<td align="right"><%=FormatNumber(ntonasepksSHI,2)%></td>
				<td align="right"><%=FormatNumber(ntonasepksSBI,2)%></td>
				<td align="right"><%=FormatNumber(nrotasiHI,2)%></td>
				<td align="right"><%=FormatNumber(nrotasiSHI,2)%></td>
				<td align="right"><%=FormatNumber(nrotasiSBI,2)%></td>
				<td align="right"><%=FormatNumber(nrestanHI,2)%></td>
				<td align="right"><%=FormatNumber(nrestanSHI,2)%></td>
				<td align="right"><%=FormatNumber(nrestanSBI,2)%></td>
				
			</tr>
</tbody>
</table>
</form>
</div>
</body>
</html>

