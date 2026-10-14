<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<html>

<head>
		<title>FBPlants - LAPORAN KEBUN</title>
		<link href="/js/tabbedpanels/SpryTabbedPanels.css" rel="stylesheet" type="text/css"/>
		<script src="/js/tabbedpanels/SpryTabbedPanels.js" type="text/javascript"></script>
		<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
		<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
		<link rel="stylesheet" href="css/jquery-ui.css" />
		<link href="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.css" rel="stylesheet" type="text/css"/>
		<script src="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.js" type="text/javascript"></script>
		<script language="JavaScript" src='/js/x.js'></script>
		<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
	</head>
<script language="JavaScript">
xAddEventListener(window, 'load',
  function() {
    new xTableHeaderFixed('xthf-site', window);
    new xTableHeaderFixed('xthf-green', window);
    new xTableHeaderFixed('xthf-blue', window);
  }, false
);
</script>
<style>
	.posisikirijudul{
		position: relative;
		left: 0px;
	}
	.posisikiri{
		position: relative;
		left: 20px;
	}
	.posisiheader{
		position: relative;
		left: 100px;
	}
	.lebarsection{
		width: 160px;
	}
	.lebargangcode{
		width: 200px;
	}
	.lebarabsen{
		width: 60px;
	}
	.lebaremployee{
		width: 230px;
	}
	.lebaractivity{
		width: 250px;
	}
	.lebaracllowedcode{
		width: 210px;
	}
	.lebarupah{
		width: 120px;
	}
	.warnagaris{
		border-color: #FFA800;
	}
	.tebal{
		font-weight: bold;
	}
	.ratakiri{
		text-align: left;
	}
	.ratakanan{
		text-align: right;
	}
	.ratatengah{
		text-align: center;
	}

	tr.top td {
	border-top: thin solid black;
	}

	tr.bottom td {
	border-bottom: thin solid black;
	}

	tr.row td:first-child {
	border-left: thin solid black;
	}

	tr.row td:last-child {
	border-right: thin solid black;
	}

	table{
    width:100%;
    border-collapse:collapse;
    table-layout:auto;
}

th, td{
    border:1px solid #999;
    padding:4px;
    font-size:14px;
}

th{
    text-align:center;
    vertical-align:middle;
}

td:first-child{
    min-width:550px;
    white-space:nowrap;
}

.ratakanan{
    text-align:right;
}

.ratakiri{
    text-align:left;
}


/* Grouping bertingkat Excel-style */
.r-hdr1   { background: #d6e4bc; }
.r-hdr2   { background: #e8f0d8; padding-left: 16px !important; }
.r-hdr3   { background: #f2f6ec; padding-left: 28px !important; }
.r-budget { background: #f9fbf5; padding-left: 40px !important; }
.r-job    { background: #ffffff; padding-left: 52px !important; font-style: italic; }

/* Warna Var%: hijau = efisien, merah = over */
td.var-ok  { color: #0f6e56; font-weight: 500; }
td.var-over{ color: #a32d2d; font-weight: 500; }

.chevron {
    display: inline-block;
    width: 16px;
    text-align: center;
    cursor: pointer;
    transition: transform 0.2s;
    font-style: normal;
}
.chevron.open { transform: rotate(90deg); }
.lbl-clickable { cursor: pointer; user-select: none; }
tr.row-hidden { display: none; }
	
<style>

/* =========================================================
   HIERARCHY SOFT PASTEL
========================================================= */

/* LEVEL 1 - Soft Blue */
.r1 td{
    background:#e8f1fd !important;
    color:#1e40af;
    font-weight:700;
    border-top:1px solid #bfdbfe;
    border-bottom:1px solid #bfdbfe;
}

/* LEVEL 2 - Soft Mint */
.r2 td{
    background:#edf9f1 !important;
    color:#166534;
    font-weight:700;
    border-top:1px solid #bbf7d0;
    border-bottom:1px solid #bbf7d0;
}

/* LEVEL 3 - Soft Cream */
.r3 td{
    background:#fff9eb !important;
    color:#92400e;
    font-weight:700;
    border-top:1px solid #fde68a;
    border-bottom:1px solid #fde68a;
}

/* LEVEL 4 - Soft Peach */
.r4 td{
    background:#fff1eb !important;
    color:#9a3412;
    font-weight:600;
    border-top:1px solid #fdba74;
    border-bottom:1px solid #fdba74;
}

/* DETAIL ROW */
.r5 td{
    background:#ffffff;
    color:#475569;
}

/* =========================================================
   HOVER
========================================================= */

tbody tr:hover td{
    background:#f8fbff !important;
    transition:0.15s;
}

/* =========================================================
   TAG WARNA LEMBUT
========================================================= */

.t1{
    background:#93c5fd;
    color:#fff;
    padding:2px 8px;
    border-radius:10px;
}

.t2{
    background:#86efac;
    color:#fff;
    padding:2px 8px;
    border-radius:10px;
}

.t3{
    background:#fde68a;
    color:#92400e;
    padding:2px 8px;
    border-radius:10px;
}





</style>


	<style>



/* =========================================================
   SUMMARY CARD
========================================================= */

.summary-wrapper{
    display:flex;
    gap:16px;
    margin-bottom:20px;
    flex-wrap:wrap;
}

.summary-card{
    flex:1;
    min-width:220px;
    background:#fff;
    border-radius:12px;
    padding:18px;
    box-shadow:0 2px 8px rgba(0,0,0,0.08);
    border-left:5px solid #0f766e;
}

.summary-title{
    font-size:13px;
    color:#64748b;
    margin-bottom:8px;
    text-transform:uppercase;
    letter-spacing:0.5px;
}

.summary-value{
    font-size:28px;
    font-weight:700;
    color:#0f172a;
}

.summary-sub{
    margin-top:8px;
    font-size:12px;
    color:#94a3b8;
}

/* =========================================================
   MODERN EXPAND ICON
========================================================= */

.chevron{
    display:inline-flex;
    align-items:center;
    justify-content:center;
    width:18px;
    height:18px;
    border-radius:50%;
    background:#e2e8f0;
    color:#334155;
    font-size:11px;
    margin-right:6px;
    transition:all .2s ease;
}

.chevron.open{
    transform:rotate(90deg);
    background:#0f766e;
    color:#fff;
}

/* =========================================================
   PROGRESS BAR
========================================================= */

.progress-wrap{
    width:100%;
    height:16px;
    background:#e5e7eb;
    border-radius:10px;
    overflow:hidden;
}

.progress-bar{
    height:100%;
    border-radius:10px;
    text-align:center;
    line-height:16px;
    font-size:10px;
    font-weight:700;
    color:#fff;
}

.p-good{
    background:#10b981;
}

.p-warning{
    background:#f59e0b;
}

.p-danger{
    background:#ef4444;
}

/* =========================================================
   NEGATIVE NUMBER
========================================================= */

.negatif{
    color:#dc2626 !important;
    font-weight:700;
}

/* =========================================================
   SCROLLBAR
========================================================= */

::-webkit-scrollbar{
    height:10px;
    width:10px;
}

::-webkit-scrollbar-track{
    background:#f1f5f9;
}

::-webkit-scrollbar-thumb{
    background:#94a3b8;
    border-radius:10px;
}

::-webkit-scrollbar-thumb:hover{
    background:#64748b;
}

/* =========================================================
   TABLE SHADOW
========================================================= */

#product-table{
    box-shadow:0 2px 8px rgba(0,0,0,0.05);
}

/* =========================================================
   HEADER SHADOW
========================================================= */

thead th{
    box-shadow:inset 0 -1px 0 rgba(255,255,255,0.15);
}

/* =========================================================
   ANIMATION
========================================================= */

tbody tr{
    transition:all .15s ease;
}



</style>


	 <%

	dim region,company,bulan,tahun,costelemen,t1,r1,fcbaname,CompanyName,MgrName,KTUName,CropName,FcbaType,coa,sqldtlupah,rsdtlupah,no
	dim tamountbl,tamountbi,tamountsbi


	dim txtcari,fcba,bkmfrom,bkmto,idview,actfrom,actto,sectioncode,statustm
	dim filterdiv,filtersource,filterfcba,filterprd,filterprdlalu,filtercostelemen,filterprdsbi,filtertmtbm
	dim prev_bulanfrom,prev_bulanto
	dim bulansplitfrom,tglsplitfrom,thnsplitfrom,tglfrom,tglto,tglawalsbi
	dim bulansplitto,tglsplitto,thnsplitto,xbln1,xthn1
	dim rphablnlalu
	dim	fisikblnlaluttl,amountblnlaluttl,rphablnlaluttl,rkb_outputblnlaluttl,rkb_amountblnlaluttl,rkbhablnlaluttl
	dim	fisikblninittl,amountblninittl,rphablninittl,rkb_outputblninittl,rkb_amountblninittl,rkbhablninittl
	dim fisiksbittl,amountsbittl,rphasbittl,rkb_outputsbittl,rkb_amountsbittl,rkbhasbittl
	dim sqlexec,periode,sqlpt,rspt,filtercompany1,fcbaloop,companyview,filtercompany,filtercompany2
	dim bulanbfr,tahunbfr,tblneraca,ptype,tamountblnini,userlogin,regionname


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
	
	
	
	if bulan = 1 then 
		bulanbfr = 12
		tahunbfr = tahun -1 
	else
		bulanbfr = bulan-1
		tahunbfr = tahun
	end if
	
	if trim(region) = "-" then 
		filtercompany1 = " AND b.region like '%' "
	else 
		filtercompany1 = " AND b.region = '" & region & "' "
	end if

	if trim(company) = "-" then 
		filtercompany1 = filtercompany1 & " AND b.fccompanycode like '%' "
	else 
		filtercompany1 = filtercompany1 & " AND b.fccompanycode = '" & company & "%' "
	end if

	if trim(fcba) = "-" then 
		filtercompany1 = filtercompany1 & " AND b.fccode like '%' "
	else 
		filtercompany1 = filtercompany1 & " AND b.fccode ='" & fcba & "' "
	end if

	sqlpt = "select b.fccode,b.fccompanyname from BUSINESSUNIT b" & _ 
	" INNER JOIN USER_COMPANY_TAB u on b.FCCOMPANYCODE=u.FCCOMPANYCODE " & _ 
	" WHERE u.FCUSERCODE='" & userlogin & "'" & _ 
	filtercompany1
	set rspt = connStr.Execute(sqlpt)
	'response.write sqlpt

	

	'filter by company
	if trim(company) = "all" then 
		filtercompany = " WHERE region = '" & region & "'"
	else
		if trim(fcba) = "-" then 
	    	filtercompany = " WHERE fccompanycode ='" & company & "' "
		else 
			filtercompany = " WHERE fccode ='" & fcba & "' "
		end if
	end if
	sqlexec = "Select region, fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit " & filtercompany 
	

	' response.write sqlexec
		set rsCat = connStr.Execute(sqlexec)
		if not rsCat.eof then
			region = rsCat("region")
			fcbaname = rsCat("fcName")
			CompanyName=rsCat("FCCOMPANYNAME")
			MgrName = rsCat("FCMANAGER")
			KTUName = rsCat("FCKTU")
			CropName=rsCat("FCCROP")
			FcbaType=rsCat("fcType")
		end if



  %>
	<body>
	<div id="head_judul" style="float:left;font-size:10px; text-align:left;">
		<h1 class="judul"></h1>				
	</div>
	<br>
	
	<div id="wrapper">
		<!--#include file="print_excel_blob.asp"-->

				<table border="0" width="100%"  class="xthf-site"  id="product-table">
					<thead>

						<tr>
							<th colspan="23"
								style="background:#FFFFFF;
									color:#000;
									border:none;
									padding:15px;
									text-align:center;">

								<div style="position:relative;">

									<div >
										
										<h4 style="text-align:left;"><%=region%></h4> 
										<h4 style="text-align:left;">PT. <%=CompanyName%></h4>
									
										<%
											IF fcbaname <> "-" then 
												fcbaname = fcba 
											ELSE 
												fcbaname= "ALL"
											END IF
										%>
										<h5 style="text-align:left;"><%=fcbaname%></h5>
									</div>

									<div>
										<h2 style="margin:0;">Realisasi vs RKB</h2>
										<h3 style="margin:0;">
											Per <%=textmonth(bulan)%>/<%=tahun%>
										</h3>
									</div>

								</div>

							</th>
						</tr>

						<tr style="background:#00A300;color:#FFF;">

							<th rowspan="3"
								style="min-width:550px;">
								Deskripsi
							</th>
							<th rowspan="3">
								UOM
							</th>

							<th colspan="7">Bulan Lalu</th>
							<th colspan="7">Bulan Ini</th>
							<th colspan="7">s/d Bulan Ini</th>

						</tr>

						<tr style="background:#00A300;color:#FFF;">

							<!-- BULAN LALU -->
							<th colspan="3">Realisasi</th>
							<th colspan="3">RKB</th>
							<th rowspan="2">Var %</th>

							<!-- BULAN INI -->
							<th colspan="3">Realisasi</th>
							<th colspan="3">RKB</th>
							<th rowspan="2">Var %</th>

							<!-- SBI -->
							<th colspan="3">Realisasi</th>
							<th colspan="3">RKB</th>
							<th rowspan="2">Var %</th>

						</tr>

						<tr style="background:#00A300;color:#FFF;">

							<!-- BL -->
							<th>Qty</th>
							<th>Rp</th>
							<th>Rp/Qty</th>

							<th>Qty</th>
							<th>Rp</th>
							<th>Rp/Qty</th>

							<!-- BI -->
							<th>Qty</th>
							<th>Rp</th>
							<th>Rp/Qty</th>

							<th>Qty</th>
							<th>Rp</th>
							<th>Rp/Qty</th>

							<!-- SBI -->
							<th>Qty</th>
							<th>Rp</th>
							<th>Rp/Qty</th>

							<th>Qty</th>
							<th>Rp</th>
							<th>Rp/Qty</th>

						</tr>

					</thead>
								
								 <% 
							

								if trim(region)="all" THEN
									filtercompany2 = " AND fcba in (select b.fccode from businessunit b INNER JOIN USER_COMPANY_TAB u on b.FCCOMPANYCODE=u.FCCOMPANYCODE where b.plantation ='FBG' AND b.isactive='TRUE') "	
								elseif trim(company)="all" THEN
									filtercompany2 = " AND fcba in (select fccode from businessunit b INNER JOIN USER_COMPANY_TAB u on b.FCCOMPANYCODE=u.FCCOMPANYCODE where b.isactive='TRUE' AND b.region ='" & region & "') "
								else
									if trim(fcba) = "-" then 
										filtercompany2 = " AND fcba in (select fccode from businessunit b INNER JOIN USER_COMPANY_TAB u on b.FCCOMPANYCODE=u.FCCOMPANYCODE where b.fccompanycode ='" & company & "') "
									else 
										filtercompany2 = " AND fcba ='" & fcba & "' "
									end if
								end if
								
								
                               sqldtlupah = "SELECT " & _
											" fchdrcode, " & _
											" CASE " & _ 
											"	WHEN NVL(fcsbhdrcode, '-') = '-' " & _
											"	AND NVL(fcssbhdrcode, '-') = '-' " & _  
											"	AND NVL(fcbudgetcode, '-') = '-' " & _
											"	THEN fchdrname ELSE ' ' " & _
											" END fchdrname, " & _
											" NVL(fcsbhdrcode, ' ') fcsbhdrcode, " & _
											" CASE " & _
											"	WHEN NVL(fcssbhdrcode, '-') = '-' " & _
											"	AND NVL(fcbudgetcode, '-') = '-'  " & _
											"	THEN NVL(fcsbhdrname, ' ') ELSE ' ' " & _
											" END fcsbhdrname, " & _
											" NVL(fcssbhdrcode, ' ') fcssbhdrcode, " & _
											" CASE " & _
											"	WHEN NVL(fcbudgetcode, '-') = '-' " & _
											"	THEN NVL(fcssbhdrname, ' ') ELSE ' ' " & _
											" END fcssbhdrname, " & _
											" NVL(fcbudgetcode, ' ') fcbudgetcode, " & _
											" CASE " & _
											"	WHEN NVL(jobcode, '-') = '-' " & _
											"	THEN NVL(fcname, ' ') ELSE ' ' " & _
											" END fcname, " & _
											" NVL(jobcode, ' ')  jobcode, " & _
											" NVL(jobname, ' ')  jobname, " & _
											" NVL(uom_unit, ' ')  uom_unit, " & _
											" CASE WHEN MAX(period_month) = 1 THEN 0 ELSE SUM(output_bl)                                                                      END AS output_bl, " & _
											" CASE WHEN MAX(period_month) = 1 THEN 0 ELSE ROUND(SUM(amount_bl), 0)                                                              END AS amount_bl, " & _
											" CASE WHEN MAX(period_month) = 1 THEN 0 ELSE ROUND(DECODE(SUM(output_bl), 0, 0, SUM(amount_bl) / SUM(output_bl)), 0)             END rpqty_bl, " & _
											" CASE WHEN MAX(period_month) = 1 THEN 0 ELSE SUM(rkb_output_bl)                                                                  END rkb_output_bl, " & _
											" CASE WHEN MAX(period_month) = 1 THEN 0 ELSE ROUND(SUM(rkb_amount_bl), 0)                                                        END rkb_amount_bl, " & _
											" CASE WHEN MAX(period_month) = 1 THEN 0 ELSE ROUND(DECODE(SUM(rkb_output_bl), 0, 0, SUM(rkb_amount_bl) / SUM(rkb_output_bl)), 0) END rkb_rpqty_bl, " & _
											" CASE WHEN MAX(period_month) = 1 THEN 0 " & _
											"	ELSE DECODE( " & _
											"		ROUND(DECODE(SUM(rkb_output_bl), 0, 0, SUM(rkb_amount_bl) / SUM(rkb_output_bl)), 0), " & _
											"		0, 0, " & _
											"		(ROUND(DECODE(SUM(output_bl), 0, 0, SUM(amount_bl) / SUM(output_bl)), 0) " & _
											"		/ ROUND(DECODE(SUM(rkb_output_bl), 0, 0, SUM(rkb_amount_bl) / SUM(rkb_output_bl)), 0)) * 100 " & _
											"	) " & _
											" END var_bl, " & _
											" SUM(output_bi)                                                                                                                    output_bi, " & _
											" ROUND(SUM(amount_bi), 0)                                                                                                         amount_bi,  " & _
											" ROUND(DECODE(SUM(output_bi), 0, 0, SUM(amount_bi) / SUM(output_bi)), 0)                                                         rpqty_bi, " & _
											" SUM(rkb_output_bi)                                                                                                               rkb_output_bi, " & _
											" ROUND(SUM(rkb_amount_bi), 0)                                                                                                     rkb_amount_bi, " & _
											" ROUND(DECODE(SUM(rkb_output_bi), 0, 0, SUM(rkb_amount_bi) / SUM(rkb_output_bi)), 0)                                             rkb_rpqty_bi, " & _
											" DECODE( " & _
											"	ROUND(DECODE(SUM(rkb_output_bi), 0, 0, SUM(rkb_amount_bi) / SUM(rkb_output_bi)), 0), " & _
											"	0, 0, " & _
											"	(ROUND(DECODE(SUM(output_bi), 0, 0, SUM(amount_bi) / SUM(output_bi)), 0) " & _
											"	/ ROUND(DECODE(SUM(rkb_output_bi), 0, 0, SUM(rkb_amount_bi) / SUM(rkb_output_bi)), 0)) * 100 " & _
											" ) var_bi, " & _
											" SUM(output_sbi)                                                                                                                  output_sbi, " & _
											" ROUND(SUM(amount_sbi), 0)                                                                                                         amount_sbi, " & _
											" ROUND(DECODE(SUM(output_sbi), 0, 0, SUM(amount_sbi) / SUM(output_sbi)), 0)                                                      rpqty_sbi, " & _
											" SUM(rkb_output_sbi)                                                                                                              rkb_output_sbi, " & _
											" ROUND(SUM(rkb_amount_sbi), 0)                                                                                                    rkb_amount_sbi, " & _
											" ROUND(DECODE(SUM(rkb_output_sbi), 0, 0, SUM(rkb_amount_sbi) / SUM(rkb_output_sbi)), 0)                                          rkb_rpqty_sbi, " & _
											" DECODE( " & _
											"	ROUND(DECODE(SUM(rkb_output_sbi), 0, 0, SUM(rkb_amount_sbi) / SUM(rkb_output_sbi)), 0), " & _
											"	0, 0, " & _
											"	(ROUND(DECODE(SUM(output_sbi), 0, 0, SUM(amount_sbi) / SUM(output_sbi)), 0) " & _
											"	/ ROUND(DECODE(SUM(rkb_output_sbi), 0, 0, SUM(rkb_amount_sbi) / SUM(rkb_output_sbi)), 0)) * 100 " & _
											" ) var_sbi " & _
										" FROM ( " & _
										"	SELECT " & _
										"		fchdrcode, fchdrname, fchdrname_eng, " & _
										"		fcsbhdrcode, fcsbhdrname, fcsbhdrname_eng, " & _
										"		fcssbhdrcode, fcssbhdrname, fcssbhdrname_eng, " & _
										"		fcbudgetcode, fcname, jobcode, jobname,uom_unit, period_month, " & _
										"		SUM(output_bl)      output_bl,   SUM(amount_bl)      amount_bl, " & _
										"		SUM(rkb_output_bl)  rkb_output_bl, SUM(rkb_amount_bl) rkb_amount_bl, " & _
										"		SUM(output_bi)      output_bi,   SUM(amount_bi)      amount_bi, " & _
										"		SUM(rkb_output_bi)  rkb_output_bi, SUM(rkb_amount_bi) rkb_amount_bi, " & _
										"		SUM(output_sbi)     output_sbi,  SUM(amount_sbi)     amount_sbi, " & _
										"		SUM(rkb_output_sbi) rkb_output_sbi, SUM(rkb_amount_sbi) rkb_amount_sbi " & _
										"	FROM ( " & _
										"		SELECT " & _
										"			b.fchdrcode, " & _
										"			b.fchdrname, " & _
										"			b.fchdrname_eng, " & _
										"			b.fcsbhdrcode, " & _
										"			b.fcsbhdrname, " & _
										"			b.fcsbhdrname_eng, " & _
										"			b.fcssbhdrcode, " & _
										"			b.fcssbhdrname, " & _
										"			b.fcssbhdrname_eng, " & _
										"			b.fccode  AS fcbudgetcode, " & _
										"			b.fcname, " & _
										"			a.fccode   AS jobcode, " & _
										"			a.fcname   AS jobname, " & _
										"			a.uom_unit   AS uom_unit, " & _
										"			5          AS period_month, " & _
										"			NVL(SUM( " & _
										"				CASE WHEN o.period_month='" & bulanbfr & "' AND o.period_year='" & tahunbfr & "' " & _
										"					THEN o.OUTPUT ELSE 0 END " & _
										"			), 0) AS output_bl, " & _
										"			NVL(SUM( " & _
										"				CASE WHEN o.period_month='" & bulanbfr & "' AND o.period_year='" & tahunbfr & "' " & _
										"					THEN o.amount ELSE 0 END " & _
										"			), 0) AS amount_bl, " & _
										"			NVL(SUM( " & _
										"				CASE WHEN o.period_month='" & bulanbfr & "' AND o.period_year='" & tahunbfr & "' " & _
										"					THEN o.rkb_output ELSE 0 END " & _
										"			), 0) AS rkb_output_bl, " & _
										"			NVL(SUM( " & _
										"				CASE WHEN o.period_month='" & bulanbfr & "' AND o.period_year='" & tahunbfr & "' " & _
										"					THEN o.rkb_amount ELSE 0 END " & _
										"			), 0) AS rkb_amount_bl, " & _
										"			NVL(SUM( " & _
										"				CASE WHEN o.period_month='" & bulan & "' AND o.period_year='" & tahun & "' " & _
										"					THEN o.OUTPUT ELSE 0 END " & _
										"			), 0) AS output_bi, " & _
										"			NVL(SUM( " & _
										"				CASE WHEN o.period_month='" & bulan & "' AND o.period_year='" & tahun & "' " & _
										"					THEN o.amount ELSE 0 END " & _
										"			), 0) AS amount_bi, " & _
										"			NVL(SUM( " & _
										"				CASE WHEN o.period_month='" & bulan & "' AND o.period_year='" & tahun & "' " & _
										"					THEN o.rkb_output ELSE 0 END " & _
										"			), 0) AS rkb_output_bi, " & _
										"			NVL(SUM( " & _
										"				CASE WHEN o.period_month='" & bulan & "' AND o.period_year='" & tahun & "' " & _
										"					THEN o.rkb_amount ELSE 0 END " & _
										"			), 0) AS rkb_amount_bi, " & _
										"			NVL(SUM( " & _
										"				CASE WHEN o.period_month<='" & bulan & "' AND o.period_year='" & tahun & "' " & _
										"					THEN o.OUTPUT ELSE 0 END " & _
										"			), 0) AS output_sbi, " & _
										"			NVL(SUM( " & _
										"				CASE WHEN o.period_month<='" & bulan & "' AND o.period_year='" & tahun & "' " & _
										"					THEN o.amount ELSE 0 END " & _
										"			), 0) AS amount_sbi, " & _
										"			NVL(SUM( " & _
										"				CASE WHEN o.period_month<='" & bulan & "' AND o.period_year='" & tahun & "' " & _
										"					THEN o.rkb_output ELSE 0 END " & _
										"			), 0) AS rkb_output_sbi, " & _
										"			NVL(SUM( " & _
										"				CASE WHEN o.period_month<='" & bulan & "' AND o.period_year='" & tahun & "' " & _
										"					THEN o.rkb_amount ELSE 0 END " & _
										"			), 0) AS rkb_amount_sbi " & _
										"		FROM jobmaster a " & _
										"		INNER JOIN jobbtmaster b " & _
										"			ON TRIM(a.job_forbudgetcode) = TRIM(b.fccode) " & _
										"		LEFT JOIN OPR_ACCOUNTREPORT_BULANAN o " & _
										"			ON TRIM(o.fccode) = TRIM(a.fccode) " & _
										"		WHERE b.FCHDRCODE=01 " & _
										filtercompany2 & _

										"		GROUP BY " & _
										"			b.fchdrcode, b.fchdrname, b.fchdrname_eng, " & _
										"			b.fcsbhdrcode, b.fcsbhdrname, b.fcsbhdrname_eng, " & _
										"			b.fcssbhdrcode, b.fcssbhdrname, b.fcssbhdrname_eng, " & _
										"			b.fccode, b.fcname, " & _
										"			a.fccode, a.fcname, a.uom_unit " & _
										"	) X " & _
										"	GROUP BY " & _
										"		fchdrcode, fchdrname, fchdrname_eng, " & _
										"		fcsbhdrcode, fcsbhdrname, fcsbhdrname_eng, " & _
										"		fcssbhdrcode, fcssbhdrname, fcssbhdrname_eng, " & _
										"		fcbudgetcode, fcname, jobcode, jobname, uom_unit, period_month " & _
										" ) CATATAN " & _
										" GROUP BY GROUPING SETS ( " & _
										"	(fchdrcode, fchdrname), " & _
										"	(fcsbhdrcode, fcsbhdrname, fchdrcode, fchdrname), " & _
										"	(fcssbhdrcode, fcssbhdrname, fcsbhdrcode, fcsbhdrname, fchdrcode, fchdrname), " & _
										"	(fcbudgetcode, fcname, fcssbhdrcode, fcssbhdrname, fcsbhdrcode, fcsbhdrname, fchdrcode, fchdrname), " & _
										"	(jobcode, jobname, uom_unit, fcbudgetcode, fcname, fcssbhdrcode, fcssbhdrname, fcsbhdrcode, fcsbhdrname, fchdrcode, fchdrname) " & _
										" ) " & _
										" HAVING SUM(amount_sbi) + SUM(rkb_amount_sbi) <> 0 " & _
										" ORDER BY fchdrcode, fcsbhdrcode, fcssbhdrcode, fcbudgetcode, jobcode" 
								'response.write sqldtlupah 'tes commit
								set rsdtlupah = connStr.Execute(sqldtlupah)
									%>
										<%
											Dim varColorBl, varColorBi, varColorSbi
											Dim rowClass, rowTag, rowTagClass, rowLabel, rowIndent
											Dim isHdr1, isHdr2, isHdr3, isBudget, isJob
											Dim hasCh, uniqueId, rowParent
											Dim clsClickable, attrOnclick, chevronHtml
											Dim fhdr, fsbhdr, fssbhdr, fbudget, fjob

											no = 0
											tamountbl=0 : tamountbi=0 : tamountsbi=0

											While Not rsdtlupah.EOF

												fhdr    = Trim(rsdtlupah("fchdrcode"))
												fsbhdr  = Trim(rsdtlupah("fcsbhdrcode"))
												fssbhdr = Trim(rsdtlupah("fcssbhdrcode"))
												fbudget = Trim(rsdtlupah("fcbudgetcode"))
												fjob    = Trim(rsdtlupah("jobcode"))

												isHdr1   = (fsbhdr = "" And fssbhdr = "" And fbudget = "" And fjob = "")
												isHdr2   = (fsbhdr <> "" And fssbhdr = "" And fbudget = "" And fjob = "")
												isHdr3   = (fssbhdr <> "" And fbudget = "" And fjob = "")
												isBudget = (fbudget <> "" And fjob = "")
												isJob    = (fjob <> "")

												If isHdr1 Then
													uniqueId   = "h_" & fhdr
													rowParent  = ""
													rowClass   = "r1"
													rowTag     = "HDR"
													rowTagClass= "t1"
													rowIndent  = 0
													rowLabel   = "<strong>" & fhdr & " " & rsdtlupah("fchdrname") & "</strong>"
													hasCh      = True
												ElseIf isHdr2 Then
													uniqueId   = "h_" & fhdr & "_" & fsbhdr
													rowParent  = "h_" & fhdr
													rowClass   = "r2"
													rowTag     = "SUB"
													rowTagClass= "t2"
													rowIndent  = 16
													rowLabel   = "<strong>" & fsbhdr & " " & rsdtlupah("fcsbhdrname") & "</strong>"
													hasCh      = True
												ElseIf isHdr3 Then
													uniqueId   = "h_" & fhdr & "_" & fsbhdr & "_" & fssbhdr
													rowParent  = "h_" & fhdr & "_" & fsbhdr
													rowClass   = "r3"
													rowTag     = "SUB2"
													rowTagClass= "t3"
													rowIndent  = 28
													rowLabel   = "<strong>" & fssbhdr & " " & rsdtlupah("fcssbhdrname") & "</strong>"
													hasCh      = True
												ElseIf isBudget Then
													uniqueId   = "b_" & fbudget
													rowParent  = "h_" & fhdr & "_" & fsbhdr & "_" & fssbhdr
													rowClass   = "r4"
													rowTag     = "BDG"
													rowTagClass= "t4"
													rowIndent  = 40
													rowLabel   = "<strong>" & fbudget & " " & rsdtlupah("fcname") & "</strong>"
													hasCh      = True
												ElseIf isJob Then
													uniqueId   = "j_" & fjob & "_" & fbudget
													rowParent  = "b_" & fbudget
													rowClass   = "r5"
													rowTag     = "JOB"
													rowTagClass= "t5"
													rowIndent  = 52
													rowLabel   = fjob & " " & rsdtlupah("jobname")
													hasCh      = False
												End If

												' ===== GANTI If() inline dengan variabel biasa =====
												If hasCh Then
													clsClickable = "lbl-clickable"
													attrOnclick  = "onclick=""toggleRow('" & uniqueId & "')"""
													chevronHtml  = "<span class=""chevron open"" id=""ch_" & uniqueId & """>&#9654;</span>"
												Else
													clsClickable = ""
													attrOnclick  = ""
													chevronHtml  = "<span style=""display:inline-block;width:16px;""></span>"
												End If

												If CDbl(rsdtlupah("var_bl"))  >= 100 Then varColorBl  = "vok" Else varColorBl  = "vno" End If
												If CDbl(rsdtlupah("var_bi"))  >= 100 Then varColorBi  = "vok" Else varColorBi  = "vno" End If
												If CDbl(rsdtlupah("var_sbi")) >= 100 Then varColorSbi = "vok" Else varColorSbi = "vno" End If
											%>
												<tr class="<%=rowClass%>"
													data-id="<%=uniqueId%>"
													data-parent="<%=rowParent%>"
													data-hasch="<%=hasCh%>">

													<td class="lbl <%=clsClickable%>"
														style="padding-left:<%=8+rowIndent%>px;white-space:nowrap;"
														<%=attrOnclick%>>
														<%=chevronHtml%>
														<span style="display:inline-block;font-size:9px;padding:1px 4px;border-radius:2px;margin-right:4px;font-weight:700;" class="<%=rowTagClass%>"><%=rowTag%></span>
														<%=rowLabel%>
														
													</td>
													<td class="num ratatengah">
														<%=rsdtlupah("uom_unit")%>
													</td>

													<!-- BL -->
													<td class="num ratakanan"><%=formatnumber(rsdtlupah("output_bl"),2)%></td>
													<td class="num ratakanan"><%=formatnumber(rsdtlupah("amount_bl"),0)%></td>
													<td class="num ratakanan"><%=formatnumber(rsdtlupah("rpqty_bl"),0)%></td>
													<td class="num ratakanan"><%=formatnumber(rsdtlupah("rkb_output_bl"),2)%></td>
													<td class="num ratakanan"><%=formatnumber(rsdtlupah("rkb_amount_bl"),0)%></td>
													<td class="num ratakanan"><%=formatnumber(rsdtlupah("rkb_rpqty_bl"),0)%></td>
													<td class="num ratakanan <%=varColorBl%>"><%=formatnumber(rsdtlupah("var_bl"),2)%></td>

													<!-- BI -->
													<td class="num ratakanan"><%=formatnumber(rsdtlupah("output_bi"),2)%></td>
													<td class="num ratakanan"><%=formatnumber(rsdtlupah("amount_bi"),0)%></td>
													<td class="num ratakanan"><%=formatnumber(rsdtlupah("rpqty_bi"),0)%></td>
													<td class="num ratakanan"><%=formatnumber(rsdtlupah("rkb_output_bi"),2)%></td>
													<td class="num ratakanan"><%=formatnumber(rsdtlupah("rkb_amount_bi"),0)%></td>
													<td class="num ratakanan"><%=formatnumber(rsdtlupah("rkb_rpqty_bi"),0)%></td>
													<td class="num ratakanan <%=varColorBi%>"><%=formatnumber(rsdtlupah("var_bi"),2)%></td>

													<!-- SBI -->
													<td class="num ratakanan"><%=formatnumber(rsdtlupah("output_sbi"),2)%></td>
													<td class="num ratakanan"><%=formatnumber(rsdtlupah("amount_sbi"),0)%></td>
													<td class="num ratakanan"><%=formatnumber(rsdtlupah("rpqty_sbi"),0)%></td>
													<td class="num ratakanan"><%=formatnumber(rsdtlupah("rkb_output_sbi"),2)%></td>
													<td class="num ratakanan"><%=formatnumber(rsdtlupah("rkb_amount_sbi"),0)%></td>
													<td class="num ratakanan"><%=formatnumber(rsdtlupah("rkb_rpqty_sbi"),0)%></td>
													<td class="num ratakanan <%=varColorSbi%> "><%=formatnumber(rsdtlupah("var_sbi"),2)%></td>
												</tr>
											<%
												no = no + 1
												If isHdr1 Then
													tamountbl  = tamountbl  - CDbl(rsdtlupah("amount_bl"))
													tamountbi  = tamountbi  - CDbl(rsdtlupah("amount_bi"))
													tamountsbi = tamountsbi - CDbl(rsdtlupah("amount_sbi"))
												End If
												rsdtlupah.MoveNext
											Wend
											rsdtlupah.close
											set rsdtlupah = Nothing
											%>
								</tr>
								
				</table>
							
	</div>
<script type="text/javascript">


// Simpan state expand: semua terbuka by default
var expandState = {};
document.querySelectorAll('tr[data-id]').forEach(function(tr){
    var id = tr.getAttribute('data-id');
    var hasCh = tr.getAttribute('data-hasch') === 'True';
    if(hasCh) expandState[id] = true;
});

function getDescendants(parentId) {
    var result = [];
    document.querySelectorAll('tr[data-parent="' + parentId + '"]').forEach(function(tr){
        result.push(tr);
        var childId = tr.getAttribute('data-id');
        getDescendants(childId).forEach(function(c){ result.push(c); });
    });
    return result;
}

function toggleRow(id) {
    expandState[id] = !expandState[id];

    // Update chevron
    var ch = document.getElementById('ch_' + id);
    if(ch) {
        if(expandState[id]) ch.classList.add('open');
        else ch.classList.remove('open');
    }

    // Sembunyikan/tampilkan semua turunan
    refreshVisibility();
}

function refreshVisibility() {
    document.querySelectorAll('tr[data-parent]').forEach(function(tr){
        var visible = isVisible(tr);
        if(visible) tr.classList.remove('row-hidden');
        else tr.classList.add('row-hidden');
    });
}

function isVisible(tr) {
    var parentId = tr.getAttribute('data-parent');
    if(!parentId || parentId === '') return true;
    // Parent harus expand
    if(!expandState[parentId]) return false;
    // Cek parent dari parent
    var parentTr = document.querySelector('tr[data-id="' + parentId + '"]');
    if(!parentTr) return true;
    return isVisible(parentTr);
}
</script>
 </body>
</html>
<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
<script src="/js/select2/select2.min.js" type="text/javascript"></script>

<script type="text/javascript" LANGUAGE="VBScript">
	new Spry.Widget.TabbedPanels("tp1", {
		defaultTab: 0
	});


function numformat(val) {
	var nStr = val.toString();
	nStr = nStr.replace(/\,/g, "");
	var x = nStr.split('.');
	var x1 = x[0];
	var x2 = x.length > 1 ? '.' + x[1] : '';
	var rgx = /(\d+)(\d{3})/;
	while (rgx.test(x1)) {
		x1 = x1.replace(rgx, '$1' + '.' + '$2');
	}
	return x1 + x2;
}
 
</script>
