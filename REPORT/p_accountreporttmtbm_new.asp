<% ValidateSession %>
<!--#include virtual="/ConndB1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<html>

<head>
<title>ePlaS Reporting - Account Report Biaya TM/TBM</title>
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
tr.subtotal-row td {
	font-weight: bold;
	background-color: #f1df1f;
	border-top: 1px solid #ccc;
}
</style>
</head>

	 <%

	dim region,company,bulan,tahun,costelemen,t1,r1,coa
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

	' ===== SUBTOTAL variables per COA =====
	dim sub_fisikblnlalu,sub_amountblnlalu,sub_rphablnlalu
	dim sub_rkb_outputblnlalu,sub_rkb_amountblnlalu,sub_rkbhablnlalu
	dim sub_fisikblnini,sub_amountblnini,sub_rphablnini
	dim sub_rkb_outputblnini,sub_rkb_amountblnini,sub_rkbhablnini
	dim sub_fisiksbittl,sub_amountsbi,sub_rphasbi
	dim sub_rkb_outputsbi,sub_rkb_amountsbi,sub_rkbhasbi
	dim prev_coa, prev_coaname

	region 		= Request("region")
	company 	= Request("company")
	fcba 		= Request("fcba")
	actfrom   	= Request("actfrom")
	actto   	= Request("actto")
	bulan   	= Request("bulan")
	tahun   	= Request("tahun")
	costelemen  = Request("costelemen")
	sectioncode	= Request("section")
	statustm	= Request("statustm")
	t1   		= Request("prd")
	coa   		= Request("coa")
	
	if trim(fcba) = "-" then 
	q = " SELECT " & _
		"  TO_CHAR(Nvl((SELECT DISTINCT start_date FROM empalldedstatus " & _
		"	  WHERE fcba in  (SELECT fccode FROM businessunit where isactive='TRUE' AND fccompanycode='" & company & "') " & _
		"	  AND YEAR  = " & tahun & "" & _
		"	  AND month = " & bulan & "" & _
		"	  ),TRUNC(To_Date('" & bulan & "/01/" & tahun & "','MM/DD/YYYY'),'MM')),'MM/DD/YYYY') PS,TO_CHAR(TRUNC(To_Date('" & bulan & "/01/" & tahun & "','MM/DD/YYYY'),'MM'),'MM/DD/YYYY') BS,  " & _
		"  TO_CHAR(Nvl((SELECT DISTINCT end_date FROM empalldedstatus " & _
		"	  WHERE fcba in  (SELECT fccode FROM businessunit where isactive='TRUE' AND fccompanycode='" & company & "') " & _
		"	  AND YEAR  = " & tahun & "" & _
		"	  AND month = " & bulan & "" & _
		"	  ),Last_day(To_Date('" & bulan & "/01/" & tahun & "','MM/DD/YYYY'))),'MM/DD/YYYY') PE,TO_CHAR(Last_day(To_Date('" & bulan & "/01/" & tahun & "','MM/DD/YYYY')),'MM/DD/YYYY') BE FROM dual " 
    else 
	
	q = " SELECT " & _
		"  TO_CHAR(Nvl((SELECT start_date FROM empalldedstatus " & _
		"	  WHERE fcba= '" & fcba & "' " & _
		"	  AND YEAR  = " & tahun & "" & _
		"	  AND month = " & bulan & "" & _
		"	  ),TRUNC(To_Date('" & bulan & "/01/" & tahun & "','MM/DD/YYYY'),'MM')),'MM/DD/YYYY') PS,TO_CHAR(TRUNC(To_Date('" & bulan & "/01/" & tahun & "','MM/DD/YYYY'),'MM'),'MM/DD/YYYY') BS,  " & _
		"  TO_CHAR(Nvl((SELECT end_date FROM empalldedstatus " & _
		"	 WHERE fcba= '" & fcba & "' " & _
		"	  AND YEAR  = " & tahun & "" & _
		"	  AND month = " & bulan & "" & _
		"	  ),Last_day(To_Date('" & bulan & "/01/" & tahun & "','MM/DD/YYYY'))),'MM/DD/YYYY') PE,TO_CHAR(Last_day(To_Date('" & bulan & "/01/" & tahun & "','MM/DD/YYYY')),'MM/DD/YYYY') BE FROM dual " 
    end if 
		' response.write q
		set rsCat = connStr.Execute(q)
		if not rsCat.eof then
		    
			if trim(t1) = "P" THEN
			actfrom = rsCat("PS")
			actto   = rsCat("PE")
			else
			actfrom = rsCat("BS")
			actto   = rsCat("BE")
			end if
		end if
		rsCat.Close
		set rsCat = Nothing
	
	if trim(fcba) = "-" then 	
	q = " SELECT " & _
		"  TO_CHAR(Nvl((SELECT Distinct start_date FROM empalldedstatus " & _
		"	  WHERE fcba in  (SELECT fccode FROM businessunit where isactive='TRUE' AND fccompanycode='" & company & "') " & _
		"	  AND YEAR  = " & tahun & "" & _
		"	  AND month = 1 " & _
		"	  ),TRUNC(To_Date('" & "01" & "/01/" & tahun & "','MM/DD/YYYY'),'MM')),'MM/DD/YYYY') PS,TO_CHAR(TRUNC(To_Date('" & "01" & "/01/" & tahun & "','MM/DD/YYYY'),'MM'),'MM/DD/YYYY') BS FROM DUAL " 
	else
	q = " SELECT " & _
		"  TO_CHAR(Nvl((SELECT start_date FROM empalldedstatus " & _
		"	  WHERE fcba= '" & fcba & "' " & _
		"	  AND YEAR  = " & tahun & "" & _
		"	  AND month = 1 " & _
		"	  ),TRUNC(To_Date('" & "01" & "/01/" & tahun & "','MM/DD/YYYY'),'MM')),'MM/DD/YYYY') PS,TO_CHAR(TRUNC(To_Date('" & "01" & "/01/" & tahun & "','MM/DD/YYYY'),'MM'),'MM/DD/YYYY') BS FROM DUAL " 
	end if
	' response.write q
		set rsCat = connStr.Execute(q)
		if not rsCat.eof then
		    
			if trim(t1) = "P" THEN
			tglawalsbi = rsCat("PS")
			else
			tglawalsbi = rsCat("BS")
			end if
		end if
		rsCat.Close
		set rsCat = Nothing	
	
		periode = actfrom & " - " & actto		

	if trim(region)="all" THEN
		filtercompany1 = " WHERE isactive='TRUE' AND b.plantation ='FBG' "	
	END IF
     if trim(company)="all" THEN
		filtercompany1 = " WHERE isactive='TRUE' AND b.region ='" & region & "' "
	else
		if trim(fcba) = "-" then 
			filtercompany1 = " WHERE isactive='TRUE' AND b.fccompanycode ='" & company & "' "
		else 
			filtercompany1 = " WHERE isactive='TRUE' AND b.fccode ='" & fcba & "' "
		end if
    end if
	
	sqlpt = "select b.fccode,b.fccompanyname from BUSINESSUNIT b" & _ 
	filtercompany1
	set rspt = connStr.Execute(sqlpt)

	if not rspt.eof then
			companyview = rspt("fccompanyname")
	''	While Not rspt.EOF
	''		fcbaloop = rspt("fccode")
	''		sqlexec = "BEGIN " & _ 
    ''          "opr_accountreport_bulanan_api.Proses('" & bulan & "','" & tahun & "','" & fcbaloop & "');" & _
	''		  "END; " 
    ''		connStr.Execute(sqlexec)

	''	rspt.MoveNext
	''	Wend
		rspt.close()
		rspt	=	null

	END IF

	q = "Select region, fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & fcba & "' "
		set rsCat = connStr.Execute(q)
		if not rsCat.eof then
			region = rsCat("region")
			fcbaname = rsCat("fcName")
			CompanyName=rsCat("FCCOMPANYNAME")
			MgrName = rsCat("FCMANAGER")
			KTUName = rsCat("FCKTU")
			CropName=rsCat("FCCROP")
			FcbaType=rsCat("fcType")
		end if
	
	bulansplitfrom  = left(actfrom,2)
	tglsplitfrom    = mid(actfrom,4,2)
	thnsplitfrom    = right(actfrom,4)
	tglfrom     	= tglsplitfrom & "/" & bulansplitfrom & "/" & thnsplitfrom

	bulansplitto  	= left(actto,2)
	tglsplitto    	= mid(actto,4,2)
	thnsplitto    	= right(actto,4)
	tglto     		= tglsplitto & "/" & bulansplitto & "/" & thnsplitto

	prev_bulanfrom  = dateadd("m", -1,tglfrom)
	prev_bulanto	= dateadd("m", -1,tglto)

	xbln1       = bulan
	xthn1       = tahun
	if  xbln1   = 1 then
		xbln2   = 12
		xthn2   = xthn1-1
	else   
		xbln2   = xbln1-1
		xthn2   = xthn1
	end if
	 
	
	if trim(region)="all" THEN
		filtercompany = " where t.fcba in (select b.fccode from businessunit b INNER JOIN USER_COMPANY_TAB u on b.FCCOMPANYCODE=u.FCCOMPANYCODE where b.plantation ='FBG' AND b.isactive='TRUE') "	
	elseif trim(company)="all" THEN
		filtercompany = " where t.fcba in (select fccode from businessunit b INNER JOIN USER_COMPANY_TAB u on b.FCCOMPANYCODE=u.FCCOMPANYCODE where b.isactive='TRUE' AND b.region ='" & region & "') "
	else
		if trim(fcba) = "-" then 
			filtercompany = " where t.fcba in (select fccode from businessunit b INNER JOIN USER_COMPANY_TAB u on b.FCCOMPANYCODE=u.FCCOMPANYCODE where b.fccompanycode ='" & company & "') "
		else 
			filtercompany = " where t.fcba ='" & fcba & "' "
		end if
	end if

	filterprdlalu = " WHERE (o.transdate between to_date('" & prev_bulanfrom & "','mm/dd/yyyy') and to_date('" & prev_bulanto & "','mm/dd/yyyy')) " 			   
	filterprd = " WHERE (o.transdate between to_date('" & actfrom & "','mm/dd/yyyy') and to_date('" & actto & "','mm/dd/yyyy')) " 			   
	filterprdsbi = " WHERE (o.transdate between to_date('" & tglawalsbi & "','mm/dd/yyyy') and to_date('" & actto & "','mm/dd/yyyy')) " 			   

	if trim(costelemen) <> "all" then 
		filtercostelemen = " AND o.source='" & costelemen & "' "
	end if

	if trim(coa) <> "" then 
		filterCoa = " AND o.fccode='" & coa & "' "
	end if

	 if trim(sectioncode) <> "all" then
	 	filterdiv = " AND o.sectioncode='" & sectioncode & "'"
	 end if

	if trim(statustm) <> "all" then
		filtertmtbm = " AND o.fccode_status='" & statustm & "'"
	end if


    sqlview =  "select t.region,t.coa,t.coaname,t.fccompanycode,t.fcba,t.fieldstatus,t.source,t.jobgroupcode,t.jobgroupname, budgetcode,budgetname," & _
	" NVL(sum(t.hkblnlalu),0)hkblnlalu,NVL(sum(t.fisikblnlalu),0)fisikblnlalu,NVL(sum(t.unitblnlalu),0)unitblnlalu,NVL(sum(t.amountblnlalu),0)amountblnlalu, " & _
	" sum(t.hkblnini)hkblnini,sum(t.fisikblnini)fisikblnini,sum(t.unitblnini)unitblnini,sum(t.amountblnini)amountblnini, " & _
	" sum(t.hksbi)hksbi,nvl(sum(t.fisiksbi),0)fisiksbi,sum(t.unitsbi)unitsbi,sum(t.amountsbi)amountsbi, " & _
	" NVL(SUM(t.rkb_hkblnlalu),0)rkb_hkblnlalu,NVL(SUM(t.rkb_outputblnlalu),0)rkb_outputblnlalu,NVL(SUM(t.rkb_amountblnlalu),0)rkb_amountblnlalu, " & _
	" NVL(SUM(t.rkb_hkblnini),0)rkb_hkblnini,NVL(SUM(t.rkb_outputblnini),0)rkb_outputblnini,NVL(SUM(t.rkb_amountblnini),0)rkb_amountblnini, " & _
	" NVL(SUM(t.rkb_hksbi),0)rkb_hksbi,NVL(SUM(t.rkb_outputsbi),0)rkb_outputsbi,NVL(SUM(t.rkb_amountsbi),0)rkb_amountsbi, " & _
	" COALESCE(NVL(SUM(t.amountblnlalu),0) / NULLIF(NVL(SUM(t.fisikblnlalu),0),0),0) rphablnlalu, " & _
	" COALESCE(NVL(SUM(t.amountblnini),0) / NULLIF(NVL(SUM(t.fisikblnini),0),0),0) rphablnini, " & _ 
	" COALESCE(NVL(SUM(t.amountsbi),0) / NULLIF(NVL(SUM(t.fisiksbi),0),0),0) rphasbi," & _
	" COALESCE(NVL(SUM(t.rkb_amountblnlalu),0) / NULLIF(NVL(SUM(t.rkb_outputblnlalu),0),0),0) rkb_rphablnlalu," & _ 
	" COALESCE(NVL(SUM(t.rkb_amountblnini),0) / NULLIF(NVL(SUM(t.rkb_outputblnini),0),0),0) rkb_rphablnini," & _ 
	" COALESCE(NVL(SUM(t.rkb_amountsbi),0) / NULLIF(NVL(SUM(t.rkb_outputsbi),0),0),0) rkb_rphasbi," & _
	" DECODE(COALESCE(NVL(SUM(t.rkb_amountblnlalu),0) / NULLIF(NVL(SUM(t.rkb_outputblnlalu),0),0),0),0,1,COALESCE(NVL(SUM(t.amountblnlalu),0) / NULLIF(NVL(SUM(t.unitblnlalu),0),0),0) / COALESCE(NVL(SUM(t.rkb_amountblnlalu),0) / NULLIF(NVL(SUM(t.rkb_outputblnlalu),0),0),0))*100 var_blnlalu, " & _ 
	" DECODE(COALESCE(NVL(SUM(t.rkb_amountblnini),0) / NULLIF(NVL(SUM(t.rkb_outputblnini),0),0),0),0,1,COALESCE(NVL(SUM(t.amountblnini),0) / NULLIF(NVL(SUM(t.unitblnini),0),0),0) / COALESCE(NVL(SUM(t.rkb_amountblnini),0) / NULLIF(NVL(SUM(t.rkb_outputblnini),0),0),0))*100 var_blnini, " & _ 
	" DECODE(COALESCE(NVL(SUM(t.rkb_amountsbi),0) / NULLIF(NVL(SUM(t.rkb_outputsbi),0),0),0),0,1,COALESCE(NVL(SUM(t.amountsbi),0) / NULLIF(NVL(SUM(t.unitsbi),0),0),0) / COALESCE(NVL(SUM(t.rkb_amountsbi),0) / NULLIF(NVL(SUM(t.rkb_outputsbi),0),0),0))*100 var_sbi " & _ 

	" from( " & _
			" SELECT b.region, o.fccode as coa, (select distinct fcname from jobmaster where fccode = o.fccode) as coaname,b.fccompanycode,b.fccode fcba,o.source,o.GROUPCODE  jobgroupcode, o.GROUPNAME  jobgroupname, o.fccode_status fieldstatus, " & _
			" o.budgetcode, o.budgetname," & _
			" decode("&bulan&",1,0,nvl(SUM(o.hk),0)) hkblnlalu, decode("&bulan&",1,0,nvl(SUM(o.OUTPUT),0)) fisikblnlalu, decode("&bulan&",1,0,nvl(SUM(o.unit),0)) unitblnlalu,decode("&bulan&",1,0,nvl(sum(o.amount),0)) amountblnlalu, " & _
			" 0 hkblnini,0 fisikblnini,0 unitblnini,0 amountblnini, " & _
			" 0 hksbi,0 fisiksbi,0 unitsbi,0 amountsbi, " & _
			" sum(o.rkb_hk) rkb_hkblnlalu,sum(o.rkb_output) rkb_outputblnlalu,sum(o.rkb_amount) rkb_amountblnlalu, " & _
			" 0 rkb_hkblnini,0 rkb_outputblnini,0 rkb_amountblnini, " & _
			" 0 rkb_hksbi,0 rkb_outputsbi,0 rkb_amountsbi " & _
			" FROM OPR_ACCOUNTREPORT_BULANAN  o " & _
			" LEFT JOIN BUSINESSUNIT b on o.FCBA=b.fccode " & _
		
			filterprdlalu & _
			filtercostelemen & _
			filtertmtbm & _
			filterdiv & _
			filterCoa & _
			" group by  b.region,o.fccode,b.fccompanycode,b.fccode,o.source,o.GROUPCODE,o.GROUPNAME,o.fccode_status,o.budgetcode,o.budgetname " & _
			"  union all " & _
			" SELECT b.region, o.fccode as coa, (select distinct fcname from jobmaster where fccode = o.fccode) as coaname,b.fccompanycode,b.fccode fcba,o.source,o.GROUPCODE  jobgroupcode, o.GROUPNAME  jobgroupname, o.fccode_status fieldstatus, " & _
			" o.budgetcode, o.budgetname," & _
			" 0 hkblnlalu,0 fisikblnlalu,0 unitblnlalu,0 amountblnlalu, " & _
			" SUM(o.hk) hkblnini,SUM(o.OUTPUT) fisikblnini,SUM(o.unit) unitblnini,sum(o.amount)amountblnini, " & _
			" 0 hksbi,0 fisiksbi,0 unitsbi,0 amountsbi, " & _
			" 0 rkb_hkblnlalu,0 rkb_outputblnlalu,0 rkb_amountblnlalu, " & _
			" SUM(o.rkb_hk)rkb_hkblnini,SUM(o.rkb_output)rkb_outputblnini,SUM(o.rkb_amount)rkb_amountblnini, " & _
			" 0 rkb_hksbi,0 rkb_outputsbi,0 rkb_amountsbi " & _
			" FROM OPR_ACCOUNTREPORT_BULANAN  o " & _
			" LEFT JOIN BUSINESSUNIT b on o.FCBA=b.fccode " & _
		
			filterprd & _
			filtercostelemen & _
			filtertmtbm & _
			filterdiv & _
			filterCoa & _
			" group by  b.region,o.fccode,b.fccompanycode,b.fccode,o.source,o.GROUPCODE,o.GROUPNAME,o.fccode_status,o.budgetcode,o.budgetname " & _
			"  union all " & _
			" SELECT b.region, o.fccode as coa, (select distinct fcname from jobmaster where fccode = o.fccode) as coaname,b.fccompanycode,b.fccode fcba,o.source,o.GROUPCODE  jobgroupcode, o.GROUPNAME  jobgroupname, o.fccode_status fieldstatus, " & _
			" o.budgetcode, o.budgetname," & _
			" 0 hkblnlalu,0 fisikblnlalu,0 unitblnlalu,0 amountblnlalu, " & _
			" 0 hkblnini,0 fisikblnini,0 unitblnini,0 amountblnini, " & _
			" SUM(o.hk) hksbi,SUM(o.OUTPUT) fisiksbi,SUM(o.unit) unitsbi,sum(o.amount)amountsbi, " & _
			" 0 rkb_hkblnlalu,0 rkb_outputblnlalu,0 rkb_amountblnlalu, " & _
			" 0 rkb_hkblnini,0 rkb_outputblnini,0 rkb_amountblnini, " & _
			" sum(o.rkb_hk) rkb_hksbi,sum(o.rkb_output) rkb_outputsbi,sum(o.rkb_amount) rkb_amountsbi " & _
			" FROM OPR_ACCOUNTREPORT_BULANAN  o " & _
			" LEFT JOIN BUSINESSUNIT b on o.FCBA=b.fccode " & _
		
			filterprdsbi & _
			filtercostelemen & _
			filtertmtbm & _
			filterdiv & _
			filterCoa & _
			" group by  b.region,o.fccode,b.fccompanycode,b.fccode,o.source,o.GROUPCODE,o.GROUPNAME,o.fccode_status,o.budgetcode,o.budgetname " & _
	" )t  " & _
	"	LEFT JOIN jobmaster a ON t.coa=a.fccode " & _
	"	INNER JOIN jobbtmaster b ON TRIM(a.job_forbudgetcode) = TRIM(b.fccode) " & _
	filtercompany & _
	" group by t.region,t.coa,t.coaname,t.fccompanycode,t.fcba,t.source,t.budgetcode,t.budgetname,t.jobgroupcode,t.jobgroupname,t.fieldstatus " & _
	" order by t.region,t.coa,t.coaname,t.fccompanycode,t.fcba,t.budgetcode,t.budgetname,t.jobgroupcode,t.jobgroupname,t.fieldstatus,t.source"
	'response.write sqlview
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
							<h4><%=region%></h4>
							<h4>PT. <%=companyview%></h4>
							<%
								IF fcbaname <> "-" then 
									fcbaname = fcba 
								ELSE 
									fcbaname= "ALL"
								END IF
							%>
							<h5><%=fcbaname%></h5>
						</div>
						<div id="subheader" class="centered" style="width:50%; margin:auto;">
							<h1 class="judul">Laporan Kebun - Rekap</h1>
							<h3>Periode: <%=periode%></h5>
						</div>
						<div style="clear:both;"></div>
					</div>
				</th>
			</tr>
			
		<tr style="background-color:#00A300;color:#fff;">
			<th colspan="3" rowspan=3>Lokasi</th>
		</tr>
		<tr style="background-color:#00A300;color:#fff;">
			<th colspan="7">Bulan Lalu</th>
			<th colspan="7">Bulan Ini</th>
			<th colspan="7">s/d Bulan Ini</th>
		</tr>
		<tr style="background-color:#00A300;color:#fff;">
			<th colspan="3">Realisasi</th>
			<th colspan="3">RKB</th>
			<th colspan="">Var</th>
			<th colspan="3">Realisasi</th>
			<th colspan="3">RKB</th>
			<th colspan="">Var</th>
			<th colspan="3">Realisasi</th>
			<th colspan="3">RKB</th>
			<th colspan="">Var</th>
		</tr>
		<tr style="background-color:#00A300;color:#fff;">
			<th>Status</th>
			<th>Pekerjaan</th>
			<th>Cost Element</th>
			<th>Output(Ha)</th>
			<th>Rupiah</th>
			<th>Rp/Ha</th>
			<th>Output(Ha)</th>
			<th>Rupiah</th>
			<th>Rp/Ha</th>
			<th>%</th>
			<th>Output(Ha)</th>
			<th>Rupiah</th>
			<th>Rp/Ha</th>
			<th>Output(Ha)</th>
			<th>Rupiah</th>
			<th>Rp/Ha</th>
			<th>%</th>
			<th>Output(Ha)</th>
			<th>Rupiah</th>
			<th>Rp/Ha</th>
			<th>Output(Ha)</th>
			<th>Rupiah</th>
			<th>Rp/Ha</th>
			<th>%</th>
		</tr>
		</thead>
		<tbody>
		<%
			if Not rsdata.eof then
				nCountofField = rsdata.fields.count-1
			end if
			
			' ===== Grand Total Variables =====
			fisikblnlaluttl		= 0
			amountblnlaluttl	= 0
			rphablnlaluttl		= 0
			rkb_outputblnlaluttl= 0
			rkb_amountblnlaluttl= 0
			rkbhablnlaluttl		= 0
			fisikblninittl		= 0
			amountblninittl		= 0
			rphablninittl		= 0
			rkb_outputblninittl	= 0
			rkb_amountblninittl	= 0
			rkbhablninittl		= 0
			fisiksbittl			= 0
			amountsbittl		= 0
			rphasbittl			= 0
			rkb_outputsbittl	= 0
			rkb_amountsbittl	= 0
			rkbhasbittl			= 0

			' ===== Subtotal Variables =====
			sub_fisikblnlalu		= 0
			sub_amountblnlalu		= 0
			sub_rphablnlalu			= 0
			sub_rkb_outputblnlalu	= 0
			sub_rkb_amountblnlalu	= 0
			sub_rkbhablnlalu		= 0
			sub_fisikblnini			= 0
			sub_amountblnini		= 0
			sub_rphablnini			= 0
			sub_rkb_outputblnini	= 0
			sub_rkb_amountblnini	= 0
			sub_rkbhablnini			= 0
			sub_fisiksbittl			= 0
			sub_amountsbi			= 0
			sub_rphasbi				= 0
			sub_rkb_outputsbi		= 0
			sub_rkb_amountsbi		= 0
			sub_rkbhasbi			= 0

			prev_coa     = ""
			prev_coaname = ""

			While Not rsdata.EOF

				' ===== Detect COA change: render subtotal row before new COA =====
				dim current_coa
				current_coa = rsdata("coa")

				if prev_coa <> "" and current_coa <> prev_coa then
					' Hitung Rp/Ha subtotal (hindari division by zero)
					dim sub_rpha_l, sub_rpha_n, sub_rpha_s, sub_rkbrpha_l, sub_rkbrpha_n, sub_rkbrpha_s
					dim sub_var_l, sub_var_n, sub_var_s

					if sub_fisikblnlalu <> 0 then
						sub_rpha_l = sub_amountblnlalu / sub_fisikblnlalu
					else
						sub_rpha_l = 0
					end if
					if sub_rkb_outputblnlalu <> 0 then
						sub_rkbrpha_l = sub_rkb_amountblnlalu / sub_rkb_outputblnlalu
					else
						sub_rkbrpha_l = 0
					end if
					if sub_rkbrpha_l <> 0 then
						sub_var_l = (sub_rpha_l / sub_rkbrpha_l) * 100
					else
						sub_var_l = 100
					end if

					if sub_fisikblnini <> 0 then
						sub_rpha_n = sub_amountblnini / sub_fisikblnini
					else
						sub_rpha_n = 0
					end if
					if sub_rkb_outputblnini <> 0 then
						sub_rkbrpha_n = sub_rkb_amountblnini / sub_rkb_outputblnini
					else
						sub_rkbrpha_n = 0
					end if
					if sub_rkbrpha_n <> 0 then
						sub_var_n = (sub_rpha_n / sub_rkbrpha_n) * 100
					else
						sub_var_n = 100
					end if

					if sub_fisiksbittl <> 0 then
						sub_rpha_s = sub_amountsbi / sub_fisiksbittl
					else
						sub_rpha_s = 0
					end if
					if sub_rkb_outputsbi <> 0 then
						sub_rkbrpha_s = sub_rkb_amountsbi / sub_rkb_outputsbi
					else
						sub_rkbrpha_s = 0
					end if
					if sub_rkbrpha_s <> 0 then
						sub_var_s = (sub_rpha_s / sub_rkbrpha_s) * 100
					else
						sub_var_s = 100
					end if
		%>
		<tr class="subtotal-row">
			<td colspan="2">SUBTOTAL</td>
			<td class="ratakanan"><%=FormatNumber(sub_fisikblnlalu)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_amountblnlalu)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_rpha_l,0)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_rkb_outputblnlalu)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_rkb_amountblnlalu)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_rkbrpha_l)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_var_l)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_fisikblnini)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_amountblnini)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_rpha_n)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_rkb_outputblnini)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_rkb_amountblnini)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_rkbrpha_n)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_var_n)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_fisiksbittl)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_amountsbi)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_rpha_s)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_rkb_outputsbi)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_rkb_amountsbi)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_rkbrpha_s)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_var_s)%></td>
		</tr>
		<%
					' Reset subtotal variables
					sub_fisikblnlalu		= 0
					sub_amountblnlalu		= 0
					sub_rphablnlalu			= 0
					sub_rkb_outputblnlalu	= 0
					sub_rkb_amountblnlalu	= 0
					sub_rkbhablnlalu		= 0
					sub_fisikblnini			= 0
					sub_amountblnini		= 0
					sub_rphablnini			= 0
					sub_rkb_outputblnini	= 0
					sub_rkb_amountblnini	= 0
					sub_rkbhablnini			= 0
					sub_fisiksbittl			= 0
					sub_amountsbi			= 0
					sub_rphasbi				= 0
					sub_rkb_outputsbi		= 0
					sub_rkb_amountsbi		= 0
					sub_rkbhasbi			= 0
				end if

				prev_coa     = current_coa
				prev_coaname = rsdata("coaname")
		%>
		<tr class='' data-fieldstatus="<%=rsdata("fieldstatus")%>" data-coa="<%=rsdata("coa")%>">
			<td>
				<%
					duaDigitAwal = Left(rsdata("coa"), 2)
					if duaDigitAwal = "64" then
				%>
					UMUM
				<%elseif duaDigitAwal = "65" then%>
					BTL
				<%else%>
					<%=rsdata("fieldstatus")%>
				<%end if%>
			</td>
			<td><%=rsdata("coa")%> - <%=rsdata("coaname")%></td>
			<td><%=rsdata("source")%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("fisikblnlalu"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("amountblnlalu"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rphablnlalu"),0)%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_outputblnlalu"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_amountblnlalu"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_rphablnlalu"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("var_blnlalu")) %></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("fisikblnini"))%></td>
			<td class="ratakanan">
				<%=FormatNumber(rsdata("amountblnini"))%><br>
				<%
					bulansplitfrom1  = left(actfrom,2)
					tglsplitfrom1    = mid(actfrom,4,2)
					thnsplitfrom1    = right(actfrom,4)
					tglfrom1     	= bulansplitfrom & "/" & tglsplitfrom1 & "/" & thnsplitfrom1

					bulansplitto1  	= left(actto,2)
					tglsplitto1    	= mid(actto,4,2)
					thnsplitto1    	= right(actto,4)
					tglto1     		= bulansplitto1 & "/" & tglsplitto1 & "/" & thnsplitto1
				%>
				<%if FormatNumber(rsdata("amountblnini")) > 0 then%>
					<%if rsdata("source") = "1.UPAH" then%>
						<a target="_blank" href="http://www.firstborneo.com/Report/p_accountreporttmtbm_new_detail.asp?region=<%=region%>&company=<%=rsdata("fccompanycode")%>&idview=1&fcba=<%=fcba%>&actfrom=undefined&actto=undefined&bulan=<%=request("bulan")%>&tahun=<%=request("tahun")%>&costelemen=1.UPAH&section=all&r1=2&prd=B&statustm=all&coa=<%=rsdata("coa")%>"
						style="color: blue">
							[Detail]
						</a>
					<%elseif rsdata("source") = "4.WIP" then %>
						<a target="_blank" href="http://www.firstborneo.com/Report/p_accountreporttmtbm_new_detail.asp?region=<%=region%>&company=<%=rsdata("fccompanycode")%>&idview=1&fcba=<%=fcba%>&actfrom=undefined&actto=undefined&bulan=<%=request("bulan")%>&tahun=<%=request("tahun")%>&costelemen=4.WIP&section=all&r1=2&prd=B&statustm=all&coa=<%=rsdata("coa")%>"
						style="color: blue">
							[Detail]
						</a>
					<%else%>
						<a target="_blank" href="http://www.firstborneo.com/p_listing_lhd_rekap.asp?pcompany=<%=rsdata("fccompanycode")%>&bkmfrom=<%=tglfrom1%>&bkmto=<%=tglto1%>&pfcba=<%=rsdata("fcba")%>&idview=26&divisi=all&jobcode=<%=rsdata("coa")%>&blok=null"
						style="color: blue">
							[Detail]
						</a>
					<%end if %>
				<%end if%>
			</td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rphablnini"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_outputblnini"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_amountblnini"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_rphablnini"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("var_blnini")) %></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("fisiksbi"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("amountsbi"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rphasbi"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_outputsbi"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_amountsbi"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_rphasbi"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("var_sbi")) %></td>
		</tr>
		<%
			' ===== Accumulate subtotal per COA =====
			sub_fisikblnlalu		= cdbl(sub_fisikblnlalu) + cdbl(rsdata("fisikblnlalu"))
			sub_amountblnlalu		= cdbl(sub_amountblnlalu) + cdbl(rsdata("amountblnlalu"))
			sub_rphablnlalu			= cdbl(sub_rphablnlalu) + cdbl(rsdata("rphablnlalu"))
			sub_rkb_outputblnlalu	= cdbl(sub_rkb_outputblnlalu) + cdbl(rsdata("rkb_outputblnlalu"))
			sub_rkb_amountblnlalu	= cdbl(sub_rkb_amountblnlalu) + cdbl(rsdata("rkb_amountblnlalu"))
			sub_rkbhablnlalu		= cdbl(sub_rkbhablnlalu) + cdbl(rsdata("rkb_rphablnlalu"))
			sub_fisikblnini			= cdbl(sub_fisikblnini) + cdbl(rsdata("fisikblnini"))
			sub_amountblnini		= cdbl(sub_amountblnini) + cdbl(rsdata("amountblnini"))
			sub_rphablnini			= cdbl(sub_rphablnini) + cdbl(rsdata("rphablnini"))
			sub_rkb_outputblnini	= cdbl(sub_rkb_outputblnini) + cdbl(rsdata("rkb_outputblnini"))
			sub_rkb_amountblnini	= cdbl(sub_rkb_amountblnini) + cdbl(rsdata("rkb_amountblnini"))
			sub_rkbhablnini			= cdbl(sub_rkbhablnini) + cdbl(rsdata("rkb_rphablnini"))
			sub_fisiksbittl			= cdbl(sub_fisiksbittl) + cdbl(rsdata("fisiksbi"))
			sub_amountsbi			= cdbl(sub_amountsbi) + cdbl(rsdata("amountsbi"))
			sub_rphasbi				= cdbl(sub_rphasbi) + cdbl(rsdata("rphasbi"))
			sub_rkb_outputsbi		= cdbl(sub_rkb_outputsbi) + cdbl(rsdata("rkb_outputsbi"))
			sub_rkb_amountsbi		= cdbl(sub_rkb_amountsbi) + cdbl(rsdata("rkb_amountsbi"))
			sub_rkbhasbi			= cdbl(sub_rkbhasbi) + cdbl(rsdata("rkb_rphasbi"))

			' ===== Accumulate grand total =====
			fisikblnlaluttl		= cdbl(fisikblnlaluttl) + cdbl(rsdata("fisikblnlalu"))
			amountblnlaluttl	= cdbl(amountblnlaluttl) + cdbl(rsdata("amountblnlalu"))
			rphablnlaluttl		= cdbl(rphablnlaluttl) + cdbl(rsdata("rphablnlalu"))
			rkb_outputblnlaluttl= cdbl(rkb_outputblnlaluttl) + cdbl(rsdata("rkb_outputblnlalu"))
			rkb_amountblnlaluttl= cdbl(rkb_amountblnlaluttl) + cdbl(rsdata("rkb_amountblnlalu"))
			rkbhablnlaluttl		= cdbl(rkbhablnlaluttl) + cdbl(rsdata("rkb_rphablnlalu"))
			fisikblninittl		= cdbl(fisikblninittl) + cdbl(rsdata("fisikblnini"))
			amountblninittl		= cdbl(amountblninittl) + cdbl(rsdata("amountblnini"))
			rphablninittl		= cdbl(rphablninittl) + cdbl(rsdata("rphablnini"))
			rkb_outputblninittl = cdbl(rkb_outputblninittl) + cdbl(rsdata("rkb_outputblnini"))
			rkb_amountblninittl = cdbl(rkb_amountblninittl) + cdbl(rsdata("rkb_amountblnini"))
			rkbhablninittl		= cdbl(rkbhablninittl) + cdbl(rsdata("rkb_rphablnini"))
			fisiksbittl			= cdbl(fisiksbittl) + cdbl(rsdata("fisiksbi"))
			amountsbittl		= cdbl(amountsbittl) + cdbl(rsdata("amountsbi"))
			rphasbittl			= cdbl(rphasbittl) + cdbl(rsdata("rphasbi"))
			rkb_outputsbittl	= cdbl(rkb_outputsbittl) + cdbl(rsdata("rkb_outputsbi"))
			rkb_amountsbittl	= cdbl(rkb_amountsbittl) + cdbl(rsdata("rkb_amountsbi"))
			rkbhasbittl			= cdbl(rkbhasbittl) + cdbl(rsdata("rkb_rphasbi"))

			rsdata.MoveNext
			Wend

			' ===== Render LAST subtotal row after loop ends =====
			if prev_coa <> "" then
				dim last_rpha_l, last_rpha_n, last_rpha_s, last_rkbrpha_l, last_rkbrpha_n, last_rkbrpha_s
				dim last_var_l, last_var_n, last_var_s

				if sub_fisikblnlalu <> 0 then
					last_rpha_l = sub_amountblnlalu / sub_fisikblnlalu
				else
					last_rpha_l = 0
				end if
				if sub_rkb_outputblnlalu <> 0 then
					last_rkbrpha_l = sub_rkb_amountblnlalu / sub_rkb_outputblnlalu
				else
					last_rkbrpha_l = 0
				end if
				if last_rkbrpha_l <> 0 then
					last_var_l = (last_rpha_l / last_rkbrpha_l) * 100
				else
					last_var_l = 100
				end if

				if sub_fisikblnini <> 0 then
					last_rpha_n = sub_amountblnini / sub_fisikblnini
				else
					last_rpha_n = 0
				end if
				if sub_rkb_outputblnini <> 0 then
					last_rkbrpha_n = sub_rkb_amountblnini / sub_rkb_outputblnini
				else
					last_rkbrpha_n = 0
				end if
				if last_rkbrpha_n <> 0 then
					last_var_n = (last_rpha_n / last_rkbrpha_n) * 100
				else
					last_var_n = 100
				end if

				if sub_fisiksbittl <> 0 then
					last_rpha_s = sub_amountsbi / sub_fisiksbittl
				else
					last_rpha_s = 0
				end if
				if sub_rkb_outputsbi <> 0 then
					last_rkbrpha_s = sub_rkb_amountsbi / sub_rkb_outputsbi
				else
					last_rkbrpha_s = 0
				end if
				if last_rkbrpha_s <> 0 then
					last_var_s = (last_rpha_s / last_rkbrpha_s) * 100
				else
					last_var_s = 100
				end if
		%>
		<tr class="subtotal-row">
			<td colspan="2">SUBTOTAL</td>
			<td class="ratakanan"><%=FormatNumber(sub_fisikblnlalu)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_amountblnlalu)%></td>
			<td class="ratakanan"><%=FormatNumber(last_rpha_l,0)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_rkb_outputblnlalu)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_rkb_amountblnlalu)%></td>
			<td class="ratakanan"><%=FormatNumber(last_rkbrpha_l)%></td>
			<td class="ratakanan"><%=FormatNumber(last_var_l)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_fisikblnini)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_amountblnini)%></td>
			<td class="ratakanan"><%=FormatNumber(last_rpha_n)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_rkb_outputblnini)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_rkb_amountblnini)%></td>
			<td class="ratakanan"><%=FormatNumber(last_rkbrpha_n)%></td>
			<td class="ratakanan"><%=FormatNumber(last_var_n)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_fisiksbittl)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_amountsbi)%></td>
			<td class="ratakanan"><%=FormatNumber(last_rpha_s)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_rkb_outputsbi)%></td>
			<td class="ratakanan"><%=FormatNumber(sub_rkb_amountsbi)%></td>
			<td class="ratakanan"><%=FormatNumber(last_rkbrpha_s)%></td>
			<td class="ratakanan"><%=FormatNumber(last_var_s)%></td>
		</tr>
		<%
			end if

			rsdata.close()
			rsdata = null
		%>
		<tr class="huruftebal " style="background-color: #f1df1f;">
			<td colspan='3'>Total</td>
			<td class="ratakanan"><span id='ttlrkpfisikblnlalu'><%=FormatNumber(fisikblnlaluttl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprpblnlalu'><%=FormatNumber(amountblnlaluttl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprphablnlalu'><%=FormatNumber(rphablnlaluttl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprkbfisikblnlalu'><%=FormatNumber(rkb_outputblnlaluttl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprkbrpblnlalu'><%=FormatNumber(rkb_amountblnlaluttl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprkbrphablnlalu'><%=FormatNumber(rkbhablnlaluttl)%></span></td>
			<td class="ratakanan"><span id='ttlrkpvarianblnlalu'>0</span></td>

			<td class="ratakanan"><span id='ttlrkpfisikblnini'><%=FormatNumber(fisikblninittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprpblnini'><%=FormatNumber(amountblninittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprphablnini'><%=FormatNumber(rphablninittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprkbfisikblnini'><%=FormatNumber(rkb_outputblninittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprkbrpblnini'><%=FormatNumber(rkb_amountblninittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprkbrphablnini'><%=FormatNumber(rkbhablninittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkpvarianblnini'>0</span></td>

			<td class="ratakanan"><span id='ttlrkpfisikblnsbi'><%=FormatNumber(fisiksbittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprpblnsbi'><%=FormatNumber(amountsbittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprphablnsbi'><%=FormatNumber(rphasbittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprkbfisikblnsbi'><%=FormatNumber(rkb_outputsbittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprkbrpblnsbi'><%=FormatNumber(rkb_amountsbittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprkbrphablnsbi'><%=FormatNumber(rkbhasbittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkpvarianblnsbi'>0</span></td>
		</tr>
	</tbody>
	</table>
	</div>
	<script type="text/javascript">
	$(document).ready(function() {
		var $table = $('#product-table');
		var $tbody = $table.find('tbody');

		// ================================================================
		// LANGKAH 1: Kumpulkan semua baris ke dalam array rows[]
		// Setiap elemen menyimpan: tipe ('data'/'subtotal'), elemen TR,
		// dan nilai fieldstatus + coa (dari data-attribute untuk data rows,
		// diwarisi dari data row sebelumnya untuk subtotal rows).
		// ================================================================
		var rows = [];
		var lastFieldstatus = '';
		var lastCoa = '';

		$tbody.children('tr').each(function() {
			var $tr = $(this);
			if ($tr.hasClass('huruftebal')) return; // skip grand total

			if ($tr.hasClass('subtotal-row')) {
				rows.push({
					type: 'subtotal',
					$tr: $tr,
					fieldstatus: lastFieldstatus,
					coa: lastCoa
				});
			} else {
				// Data row: baca fieldstatus & coa dari data-attribute
				lastFieldstatus = $tr.data('fieldstatus') || $tr.find('td').eq(0).text().trim();
				lastCoa         = $tr.data('coa')         || '';
				rows.push({
					type: 'data',
					$tr: $tr,
					fieldstatus: lastFieldstatus,
					coa: lastCoa
				});
			}
		});

		// ================================================================
		// LANGKAH 2: Merge kolom 1 & 2 (Pekerjaan, Cost Element)
		// Hanya pada data rows, dalam scope fieldstatus yang sama.
		// ================================================================
		var dataRows = rows.filter(function(r){ return r.type === 'data'; });

		// Kolom 1 (Pekerjaan): merge bila coa sama DAN fieldstatus sama
		(function() {
			var prevCell = null, prevCoa = null, prevFs = null, span = 1;
			dataRows.forEach(function(r) {
				var $cell = r.$tr.find('td').eq(1);
				if (prevCell && r.coa === prevCoa && r.fieldstatus === prevFs) {
					span++;
					prevCell.attr('rowspan', span);
					$cell.hide();
				} else {
					prevCell = $cell.attr('rowspan', 1).show();
					prevCoa  = r.coa;
					prevFs   = r.fieldstatus;
					span     = 1;
				}
			});
		})();

		// Kolom 2 (Cost Element): merge bila source sama DAN coa sama DAN fieldstatus sama
		(function() {
			var prevCell = null, prevSrc = null, prevCoa = null, prevFs = null, span = 1;
			dataRows.forEach(function(r) {
				var $cell  = r.$tr.find('td').eq(2);
				var srcVal = $cell.text().trim();
				if (prevCell && srcVal === prevSrc && r.coa === prevCoa && r.fieldstatus === prevFs) {
					span++;
					prevCell.attr('rowspan', span);
					$cell.hide();
				} else {
					prevCell = $cell.attr('rowspan', 1).show();
					prevSrc  = srcVal;
					prevCoa  = r.coa;
					prevFs   = r.fieldstatus;
					span     = 1;
				}
			});
		})();

		// ================================================================
		// LANGKAH 3: Merge kolom 0 (Status/BIBITAN)
		// Proses SEMUA rows (data + subtotal) berdasarkan fieldstatus.
		// Subtotal row sudah mewarisi fieldstatus dari data row di atasnya,
		// sehingga otomatis masuk ke dalam group yang sama.
		// Subtotal row TIDAK punya cell kolom 0 (colspan="2" di HTML),
		// jadi kita hanya perlu extend rowspan dari cell data row pertama
		// dalam group — tanpa perlu menyembunyikan apapun di subtotal row.
		// ================================================================
		(function() {
			var prevCell = null, prevFs = null, span = 1;
			rows.forEach(function(r) {
				if (r.type === 'data') {
					var $cell = r.$tr.find('td').eq(0);
					if (prevCell && r.fieldstatus === prevFs) {
						// Masih group yang sama: extend rowspan cell kepala group
						span++;
						prevCell.attr('rowspan', span);
						$cell.hide();
					} else {
						// Group baru
						prevCell = $cell.attr('rowspan', 1).show();
						prevFs   = r.fieldstatus;
						span     = 1;
					}
				} else {
					// Subtotal row: tidak ada cell kolom 0 (colspan="2"),
					// cukup tambah 1 ke rowspan cell kepala group saat ini
					if (prevCell && r.fieldstatus === prevFs) {
						span++;
						prevCell.attr('rowspan', span);
					}
					// Jika subtotal muncul tanpa data row sebelumnya (edge case),
					// tidak ada yang perlu dilakukan
				}
			});
		})();
	});
	</script>
 </tbody>	

<!--#include file="bottom.Asp"-->
