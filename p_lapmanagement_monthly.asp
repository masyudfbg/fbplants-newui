<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<html>

<head>
		<title>FBPlants - LAPORAN MANAGEMENT</title>
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
		<!-- Tampilan baru (kartu judul, tab, tabel Realisasi Budget); data & Export Excel tidak berubah -->
		<meta name="viewport" content="width=device-width, initial-scale=1">
		<link rel="stylesheet" href="/css/fbp-lapmgmt.css?v=20261010a" type="text/css" />
		<script src="/js/fbp-lapmgmt.js?v=20261010a" defer></script>
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
	
</style>

	 <%

	dim region,company,bulan,tahun,costelemen,t1,r1,fcbaname,CompanyName,MgrName,KTUName,CropName,FcbaType
	dim txtcari,fcba,bkmfrom,bkmto,idview,actfrom,actto,sectioncode,statustm
	dim filterdiv,filtersource,filterfcba,filterprd,filterprdlalu,filtercostelemen,filterprdsbi,filtertmtbm
	dim prev_bulanfrom,prev_bulanto
	dim bulansplitfrom,tglsplitfrom,thnsplitfrom,tglfrom,tglto,tglawalsbi
	dim bulansplitto,tglsplitto,thnsplitto,xbln1,xthn1
	dim rphablnlalu
	dim	fisikblnlaluttl,amountblnlaluttl,rphablnlaluttl,rkb_outputblnlaluttl,rkb_amountblnlaluttl,rkbhablnlaluttl
	dim	fisikblninittl,amountblninittl,rphablninittl,rkb_outputblninittl,rkb_amountblninittl,rkbhablninittl
	dim fisiksbittl,amountsbittl,rphasbittl,rkb_outputsbittl,rkb_amountsbittl,rkbhasbittl
	dim sqlexec,periode,sqlpt,rspt,filtercompany1,fcbaloop,companyview,filtercompany
	dim bulanbfr,tahunbfr,tblneraca,ptype,tamountblnini,userlogin

	region 		= Request("region")
	company 	= Request("company")
	fcba 		= Request("fcba")
	bulan   	= Request("bulan")
	tahun   	= Request("tahun")
	ptype       = Request("ptype")
	userlogin   = session("CodeName")
	
	if ptype = 12 then
		tblneraca = "fac_trialbalance"
	else
		tblneraca = "fac_trialbalance_spt_adt"
	end if
	
	if bulan = 1 then 
		'bulanbfr = 12
		bulanbfr = ptype
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

	if not rspt.eof then

		companyview = rspt("fccompanyname")
		While Not rspt.EOF
			fcbaloop = rspt("fccode")
			sqlexec = "BEGIN " & _ 
			  "fac_laporan_management_api.Proses('" & bulan & "','" & tahun & "','" & fcbaloop & "');" & _
			  "END; " 
    		connStr.Execute(sqlexec)

		rspt.MoveNext
		Wend
		rspt.close()
		rspt	=	null

	END IF

	

	

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
	<table width="100%" id="product-table" class="xthf-site" id="table1">
	<tr>
		<th align="left" colspan="23" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">LAPORAN MANAGEMENT</th>
	</tr>
	<tr>
		<th align="left" colspan="23" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;"><%=region%>-<%=company%>-<%=fcba%></th>
	</tr>
	
	<tr>
			<td>
				<div class="TabbedPanels" id="tp1" align='center'>
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">Neraca</li>
						<li class="TabbedPanelsTab">Laba (Rugi)</li>
						<li class="TabbedPanelsTab">Arus Kas</li>
						<li class="TabbedPanelsTab">Catatan Laporan</li>
						<li class="TabbedPanelsTab">Laporan Produksi</li>
						<li class="TabbedPanelsTab">Realisasi Budget</li>
						
						
					</ul>
					<div class="TabbedPanelsContentGroup">
	<!-- Neraca -->
						<div class="TabbedPanelsContent">
							<table border="0" width="100%"  class='data' id='table2'>
								<tr >									
									<td style="border:none"><strong>PT. <%=CompanyName%></strong></td>
									<td style="border:none" colspan=3></td>
									<td style="border:none" align=right><strong>PT. <%=CompanyName%></strong></td>
								</tr>
								<tr id="trheader2">									
									<td style="border:none"><strong>NERACA</strong></td>
									<td style="border:none" colspan=3></td>
									<td style="border:none" align=right><strong>BALANCE SHEET</strong></td>
								</tr>
								<tr id="trheader2">									
									<td style="border:none">Per <%=textmonth(bulan)%>/<%=tahun%></td>
									<td style="border:none" colspan=3></td>
									<td style="border:none" align=right>As <%=textmonth(bulan)%>/<%=tahun%></td>
								</tr>
								<tr id="trheader2">									
									<td style="border:none">(dinyatakan dalam Rupiah)</td>
									<td style="border:none" colspan=3></td>
									<td style="border:none" align=right>(Stated In Rupiah)</td>
								</tr>
								<tr id="trheader2">									
									<td class="ratatengah">Deskripsi</td> <!-- head -->
									<td class="ratatengah"><%=textmonth(bulan)%>/<%=tahun%></td>
									<td class="ratatengah">Catatan/Notes</td>
									<td class="ratatengah"><%=textmonth(ptype)%>/<%=tahun-1%></td>
									<td class="ratatengah">Description</td>
								</tr>
								<% 
								'get data detail upah
								dim sqldtlupah,rsdtlupah,no,totalamount,oRs

								'if trim(fcba) = "-" then 
								'	filtercompany1 = " AND fcba in (select fccode from businessunit where fccompanycode ='" & company & "') "
								'else 
								'	filtercompany1 = " AND fcba ='" & fcba & "' "
								'end if
								
								filtercompany1 = " AND fcba in (select fccode from businessunit b" & _ 
								" INNER JOIN USER_COMPANY_TAB u on b.FCCOMPANYCODE=u.FCCOMPANYCODE " & _ 
								" WHERE u.FCUSERCODE='" & userlogin & "' " 
								if trim(region) = "-" then 
									filtercompany1 = filtercompany1 & " AND b.region like '%' "
								else 
									filtercompany1 = filtercompany1 & " AND b.region = '" & region & "' "
								end if

								if trim(company) = "-" then 
									filtercompany1 = filtercompany1 & " AND b.fccompanycode like '%' "
								else 
									filtercompany1 = filtercompany1 & " AND b.fccompanycode = '" & company & "' "
								end if

								if trim(fcba) = "-" then 
									filtercompany1 = filtercompany1 & " AND b.fccode like '%' )"
								else 
									filtercompany1 = filtercompany1 & " AND b.fccode ='" & fcba & "' )"
								end if

								'sqldtlupah = "SELECT row_id,description,sum(current_amount)current_amount,nonotes,sum(begin_amount)begin_amount,description2,rowtype " & _ 
								'			" FROM fac_laporan_management a " & _
								'			" where a.reportcode='LM_BS'" & _ 
								'			" AND period_month=" & bulan & " AND period_year=" & tahun & " " & _ 
                                '           filtercompany1 & _ 
								'			" GROUP BY row_id,description,nonotes,description2,rowtype" & _
								'			" ORDER BY row_id "
											
								sqldtlupah = " SELECT " & _
											 "  fchdrcode, " & _
											 "  CASE WHEN  Nvl(fcshdrcode,'-')='-' AND  Nvl(fcbudgetcode,'-') ='-'  THEN fchdrname " & _
											 " 						ELSE ' ' " & _
											 " 						END  fchdrname, " & _ 										
											 "  Nvl(fcshdrcode,' ')  fcshdrcode, " & _
											 "  CASE WHEN  Nvl(fcbudgetcode,'-') ='-'  THEN Nvl(fcshdrname,' ') " & _
											 " 						ELSE ' ' " & _
											 " 						END  fcshdrname, " & _ 										
										     "  Nvl(fcbudgetcode,' ') fcbudgetcode, " & _
											 "  Nvl(parentcode,' ') parentcode," & _
											 "  Nvl(parentdesc,' ')parentdesc, " & _											
										     "  Round(Sum(amount_bi),0) amount_bi, " & _
											 "  Round(Sum(amount_bl),0) amount_bl, " & _
											 "  CASE WHEN  Nvl(fcshdrcode,'-')='-' AND  Nvl(fcbudgetcode,'-') ='-'  THEN fchdrname_eng " & _
											 "			  WHEN  Nvl(fcbudgetcode,'-') ='-' THEN  fcshdrname_eng " & _
											 "	ELSE  parentdesc_eng " & _
											 " 	END  eng " & _										
											 "  FROM " & _
											 " ( " & _
											 "  	SELECT   " & _
											 "  	fchdrcode, fchdrname,fchdrname_eng, " & _
											 " 		fcshdrcode, fcshdrname,fcshdrname_eng, " & _
											 "      parentcode, parentdesc, parentdesc_eng," & _
											 "      fcbudgetcode," & _
											 " 	    Sum(amount_bl) amount_bl, " & _
											 " 	    Sum(amount_bi) amount_bi " & _
											 " 	    FROM " & _
											 "  	( " & _
											 " 		SELECT  " & _												
											 " 		fchdrcode, fchdrname,fchdrname_eng, " & _
											 " 		fcshdrcode, fcshdrname,fcshdrname_eng, " & _
											 " 		b.parentcode, b.parentdesc,b.parentdesc_eng, " & _                    
											 " 		fchdrcode||fcshdrcode||b.parentcode fcbudgetcode, " & _										 	
											 " 	    Nvl(amount_bl,0) amount_bl, " & _
											 " 		Nvl(amount_bi,0) amount_bi  " & _
											 " 		FROM " & _
											 " 		jobmaster a " & _
											 " 		INNER JOIN jobgroupmaster b ON (Trim(a.jobgroupcode)=Trim(b.fccode)) " & _
											 " 		LEFT JOIN " & _
											 " 		( " & _
											 " 		  SELECT period_year,period_month,fcba,fccode, " & _
											 " 		  CASE WHEN SubStr(fccode,1,1) = '1' then ending_balance  ELSE ending_balance*-1 END  amount_bl,0 amount_bi " & _                                                                                    
											 " 		  FROM  " & tblneraca & "" & _
											 " 		  WHERE 1=1 " & _
											 " 		  AND period_month= " & ptype & " AND period_year= " & tahun-1 & "  " & _
											            filtercompany1 & _ 
											 " 		  AND ending_balance <> 0 " & _
											 " 		  UNION ALL " & _
											 " 		  SELECT period_year,period_month,fcba,fccode, " & _
											 " 		  0 amount_bl,CASE WHEN SubStr(fccode,1,1) = '1' then ending_balance  ELSE ending_balance*-1 END  amount_bi  " & _                                                                                    
											 " 		  FROM  " & tblneraca & " " & _
											 " 		  WHERE 1=1 " & _
											 " 		  AND period_month= " & bulan & "   AND period_year= " & tahun & "  " & _
											            filtercompany1 & _
											 " 	      AND ending_balance <> 0 " & _
											 " 		)c ON (a.fccode=c.fccode) " & _
											 " 	)X " & _
											 " 	GROUP BY  fchdrcode, fchdrname,fchdrname_eng,fcshdrcode, fcshdrname,fcshdrname_eng,parentcode, parentdesc, parentdesc_eng, fcbudgetcode " & _
											 " )CATATAN " & _
											 " GROUP BY " & _
											 " GROUPING SETS ( " & _
											 " 					(fchdrcode,fchdrname,fchdrname_eng), " & _
											 " 					(fcshdrcode,fcshdrname,fcshdrname_eng,fchdrcode,fchdrname,fchdrname_eng), " & _
											 " 					(fcbudgetcode,parentcode, parentdesc,parentdesc_eng,fcshdrcode,fcshdrname,fchdrcode,fchdrname) " & _
											 "   				 ) " & _
											 " ORDER BY fchdrcode,fcshdrcode,fcbudgetcode,parentcode "

    

								set rsdtlupah = connStr.Execute(sqldtlupah)
								Dim tamountbl,tamountbi,tamountsbi,tawalkas_bi,tawalkas_bl,takhirkas_bi,takhirkas_bl
								no = 0
								tamountbl=0
								tamountbi=0
								tamountsbi=0
								tawalkas_bi=0
								tawalkas_bl=0
								takhirkas_bi=0
								takhirkas_bl=0
								
								
								While Not rsdtlupah.EOF %>
								<tr id="tru<%=no%>" >
									<td class="ratakiri"><strong><%=rsdtlupah("fchdrname")%>&nbsp;&nbsp;&nbsp;<%=rsdtlupah("fcshdrname")%></strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsdtlupah("parentdesc")%></td> <!-- head -->
									<td class="ratakanan"><%=formatnumber(rsdtlupah("amount_bi"),2)%></td>
									<td class="ratatengah"><%=rsdtlupah("parentcode")%></td>
									<td class="ratakanan"><%=formatnumber(rsdtlupah("amount_bl"),2)%></td>
									<td class="ratakanan"><%if rsdtlupah("parentcode")=" " then %><strong><%=rsdtlupah("eng")%></strong><%else%><%=rsdtlupah("eng")%><%end if%></td>
								<%
									no = no +1
									if rsdtlupah("parentcode")="01" then tawalkas_bi = cdbl(rsdtlupah("amount_bl")) end if
									rsdtlupah.MoveNext
									Wend
									rsdtlupah.close 
									set rsdtlupah = Nothing
								%>
								</tr>
								
							</table>
								<br>
							<br>
Lihat catatan atas laporan keuangan yang merupakan bagian yang tidak terpisahkan dari laporan keuangan secara keseluruhan<br>
See the accompanying notes to financial statements which are an integral part of the financial statements
						</div>
	<!-- Laba (Rugi) -->
						<div class="TabbedPanelsContent">
							<table border="0" width="100%"  class='data' id='table2'>
								<tr >									
									<td style="border:none"><strong>PT. <%=CompanyName%></strong></td>
									<td style="border:none" colspan=4></td>
									<td style="border:none" align=right><strong>PT. <%=CompanyName%></strong></td>
								</tr>
								<tr id="trheader2">									
									<td style="border:none"><strong>LABA(RUGI)</strong></td>
									<td style="border:none" colspan=4></td>
									<td style="border:none" align=right><strong>PROFIT (LOSS)</strong></td>
								</tr>
								<tr id="trheader2">									
									<td style="border:none">Per <%=textmonth(bulan)%>/<%=tahun%></td>
									<td style="border:none" colspan=4></td>
									<td style="border:none" align=right>As <%=textmonth(bulan)%>/<%=tahun%></td>
								</tr>
								<tr id="trheader2">									
									<td style="border:none">(dinyatakan dalam Rupiah)</td>
									<td style="border:none" colspan=4></td>
									<td style="border:none" align=right>(Stated In Rupiah)</td>
								</tr>
								<tr id="trheader2">									
									<td class="ratatengah">Deskripsi</td> <!-- head -->
									<td class="ratatengah">Bulan Ini</td> <!-- head -->
									<td class="ratatengah">s.d <%=textmonth(bulan)%>/<%=tahun%></td>
									<td class="ratatengah">Catatan/Notes</td>
									<td class="ratatengah"><%=textmonth(ptype)%>/<%=tahun-1%></td>
									<td class="ratatengah">Description</td>
								</tr>
								<% 
								
								
								filtercompany1 = " AND fcba in (select fccode from businessunit b" & _ 
								" INNER JOIN USER_COMPANY_TAB u on b.FCCOMPANYCODE=u.FCCOMPANYCODE " & _ 
								" WHERE u.FCUSERCODE='" & userlogin & "' " 
								
								if trim(region) = "-" then 
									filtercompany1 = filtercompany1 & " AND b.region like '%' "
								else 
									filtercompany1 = filtercompany1 & " AND b.region = '" & region & "' "
								end if

								if trim(company) = "-" then 
									filtercompany1 = filtercompany1 & " AND b.fccompanycode like '%' "
								else 
									filtercompany1 = filtercompany1 & " AND b.fccompanycode = '" & company & "' "
								end if

								if trim(fcba) = "-" then 
									filtercompany1 = filtercompany1 & " AND b.fccode like '%' )"
								else 
									filtercompany1 = filtercompany1 & " AND b.fccode ='" & fcba & "' )"
								end if

								
								sqldtlupah = " SELECT " & _
											 "  fchdrcode, " & _
											 "  CASE WHEN  Nvl(fcbudgetcode,'-') ='-'  THEN fchdrname " & _
											 " 	ELSE ' ' " & _
											 " 	END  fchdrname, " & _ 										
											 "  Nvl(fcbudgetcode,' ') fcbudgetcode, " & _
											 "  Nvl(parentcode,' ') parentcode," & _
											 "  Nvl(parentdesc,' ')parentdesc, " & _											
										     "  Round(Sum(amount_bi),0) amount_bi, " & _
											 "  Round(Sum(amount_bl),0) amount_bl, " & _
											 "  Round(Sum(amount_blnini),0) amount_blnini, " & _
											 "  CASE WHEN  Nvl(fcbudgetcode,'-') ='-'  THEN fchdrname_eng " & _
											 "	ELSE  parentdesc_eng " & _
											 " 	END  eng " & _										
											 "  FROM " & _
											 " ( " & _
											 "  	SELECT   " & _
											 "  	fchdrcode, fchdrname,fchdrname_eng, " & _
											 " 		fcshdrcode, fcshdrname,fcshdrname_eng, " & _
											 "      parentcode, parentdesc, parentdesc_eng," & _
											 "      fcbudgetcode," & _
											 " 	    Sum(amount_bl) amount_bl, " & _
											 " 	    Sum(amount_bi) amount_bi, " & _
											 " 	    Sum(amount_blnini) amount_blnini " & _
											 " 	    FROM " & _
											 "  	( " & _
											 " 		SELECT  " & _												
											 " 		fchdrcode, fchdrname,fchdrname_eng, " & _
											 " 		fcshdrcode, fcshdrname,fcshdrname_eng, " & _
											 " 		b.fccode parentcode, b.fcname parentdesc,b.fcname_eng parentdesc_eng, " & _                    
											 " 		fchdrcode||fcshdrcode||b.fccode fcbudgetcode, " & _										 	
											 " 	    Nvl(amount_bl,0) amount_bl, " & _
											 " 		Nvl(amount_bi,0) amount_bi,  " & _
											 " 		Nvl(amount_blnini,0) amount_blnini  " & _
											 " 		FROM " & _
											 " 		jobmaster a " & _
											 " 		INNER JOIN jobismaster b ON (Trim(a.job_forprofitandlosscode)=Trim(b.fccode)) " & _
											 " 		LEFT JOIN " & _
											 " 		( " & _
											 " 		  SELECT period_year,period_month,fcba,fccode, " & _
											 " 		  CASE WHEN SubStr(fccode,1,1) = '1' then ending_balance  ELSE ending_balance*-1 END  amount_bl,0 amount_bi,0 amount_blnini " & _                                                                                    
											 " 		  FROM  " & tblneraca &  "" & _
											 " 		  WHERE 1=1 " & _
											 " 		  AND period_month= " & ptype & "  AND period_year= " & tahun-1 & "  " & _
											            filtercompany1 & _ 
											 " 		  AND ending_balance <> 0 " & _
											 " 		  UNION ALL " & _
											 " 		  SELECT period_year,period_month,fcba,fccode, " & _
											 " 		  0 amount_bl,CASE WHEN SubStr(fccode,1,1) = '1' then ending_balance  ELSE ending_balance*-1 END  amount_bi,0 amount_blnini  " & _                                                                                    
											 " 		  FROM  " & tblneraca & " " & _
											 " 		  WHERE 1=1 " & _
											 " 		  AND period_month= " & bulan & "   AND period_year= " & tahun & "  " & _
											            filtercompany1 & _
											 " 	      AND ending_balance <> 0 " & _
											 "		  UNION ALL " & _
											 "		SELECT period_year,period_month,fcba,fccode, 0 amount_bl,0 amount_bi,(debet+credit)*-1 amount_blnini " & _
											 "		FROM " & tblneraca & "  " & _
											 "		WHERE 1=1 AND period_month= " & bulan & "  AND period_year= " & tahun & "  " & _
											 		filtercompany1 & _ 
											 "		AND ending_balance <> 0) c ON (a.fccode=c.fccode) " & _
											 " 	) X " & _
											 " 	GROUP BY  fchdrcode, fchdrname,fchdrname_eng,fcshdrcode, fcshdrname,fcshdrname_eng,parentcode, parentdesc, parentdesc_eng, fcbudgetcode " & _
											 " )CATATAN " & _
											 " GROUP BY " & _
											 " GROUPING SETS ( " & _
											 " 					(fchdrcode,fchdrname,fchdrname_eng), " & _
											 " 					(fcbudgetcode,parentcode, parentdesc,parentdesc_eng,fcshdrcode,fcshdrname,fchdrcode,fchdrname) " & _
											 "   				 ) " & _
											 " ORDER BY fchdrcode,fcshdrcode,fcbudgetcode,parentcode "

							'	Response.Write sqldtlupah
								set rsdtlupah = connStr.Execute(sqldtlupah)
								no = 0
								tamountbl=0
								tamountbi=0
								tamountblnini=0
								
								While Not rsdtlupah.EOF %>
								<tr id="tru<%=no%>" >
									<td class="ratakiri"><strong><%=rsdtlupah("fchdrname")%>&nbsp;&nbsp;&nbsp;</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsdtlupah("parentdesc")%></td> <!-- head -->
									<td class="ratakanan"><%=formatnumber(rsdtlupah("amount_blnini"),2)%></td>
									<td class="ratakanan"><%=formatnumber(rsdtlupah("amount_bi"),2)%></td>
									<td class="ratakiri"></td>
									<td class="ratakanan"><%=formatnumber(rsdtlupah("amount_bl"),2)%></td>
									<td class="ratakanan"><%if rsdtlupah("parentcode")=" " then %><strong><%=rsdtlupah("eng")%></strong><%else%><%=rsdtlupah("eng")%><%end if%></td>
																<%
									no = no +1
									if rsdtlupah("parentcode")=" " THEN
										tamountblnini = tamountblnini + cdbl(rsdtlupah("amount_blnini"))										
										tamountbl = tamountbl + cdbl(rsdtlupah("amount_bl"))
										tamountbi = tamountbi + cdbl(rsdtlupah("amount_bi"))										
									end if
									rsdtlupah.MoveNext
									Wend
									rsdtlupah.close 
									set rsdtlupah = Nothing
								%>
								</tr>
								<tr id="tru<%=no%>" >
									<td class="ratakiri"><strong>LABA (RUGI) BERJALAN</strong></td> <!-- head -->
									<td class="ratakanan"><%=formatnumber(tamountblnini,2)%></td>
									<td class="ratakanan"><%=formatnumber(tamountbi,2)%></td>
									<td class="ratakiri"></td>
									<td class="ratakanan"><%=formatnumber(tamountbl,2)%></td>
									<td class="ratakanan"><strong>PROFIT (LOSS) CURRENT YEAR</strong></td>
								</tr>	
								
							</table>
							<br>
							<br>
							Lihat catatan atas laporan keuangan yang merupakan bagian yang tidak terpisahkan dari laporan keuangan secara keseluruhan<br>
See the accompanying notes to financial statements which are an integral part of the financial statements

						</div>
	<!-- Arus Kas -->
						<div class="TabbedPanelsContent">
							<table border="0" width="100%"  class='data' id='table2'>
								<tr >									
									<td style="border:none"><strong>PT. <%=CompanyName%></strong></td>
									<td style="border:none" colspan=3></td>
									<td style="border:none" align=right><strong>PT. <%=CompanyName%></strong></td>
								</tr>
								<tr id="trheader2">									
									<td style="border:none"><strong>ARUS KAS</strong></td>
									<td style="border:none" colspan=3></td>
									<td style="border:none" align=right><strong>CASH FLOW</strong></td>
								</tr>
								<tr id="trheader2">									
									<td style="border:none">Per <%=textmonth(bulan)%>/<%=tahun%></td>
									<td style="border:none" colspan=3></td>
									<td style="border:none" align=right>As <%=textmonth(bulan)%>/<%=tahun%></td>
								</tr>
								<tr id="trheader2">									
									<td style="border:none">(dinyatakan dalam Rupiah)</td>
									<td style="border:none" colspan=3></td>
									<td style="border:none" align=right>(Stated In Rupiah)</td>
								</tr>
								<tr id="trheader2">									
									<td class="ratatengah">Deskripsi</td> <!-- head -->
									<td class="ratatengah"><%=textmonth(bulan)%>/<%=tahun%></td>
									<td class="ratatengah">Catatan/Notes</td>
									<td class="ratatengah"><%=textmonth(ptype)%>/<%=tahun-1%></td>
									<td class="ratatengah">Description</td>
								</tr>
								<% 
								'if trim(fcba) = "-" then 
								'	filtercompany1 = " AND fcba in (select fccode from businessunit where fccompanycode ='" & company & "') "
								'else 
								'	filtercompany1 = " AND fcba ='" & fcba & "' "
								'end if
								
								filtercompany1 = " AND fcba in (select fccode from businessunit b" & _ 
								" INNER JOIN USER_COMPANY_TAB u on b.FCCOMPANYCODE=u.FCCOMPANYCODE " & _ 
								" WHERE u.FCUSERCODE='" & userlogin & "' " 
								if trim(region) = "-" then 
									filtercompany1 = filtercompany1 & " AND b.region like '%' "
								else 
									filtercompany1 = filtercompany1 & " AND b.region = '" & region & "' "
								end if

								if trim(company) = "-" then 
									filtercompany1 = filtercompany1 & " AND b.fccompanycode like '%' "
								else 
									filtercompany1 = filtercompany1 & " AND b.fccompanycode = '" & company & "' "
								end if

								if trim(fcba) = "-" then 
									filtercompany1 = filtercompany1 & " AND b.fccode like '%' )"
								else 
									filtercompany1 = filtercompany1 & " AND b.fccode ='" & fcba & "' )"
								end if
								
								'cari saldo awal kas BI
								sqldtlupah = "SELECT NVL(Sum(ending_balance),0) sawal FROM " & _
											 "" & tblneraca & " a " & _
											 "INNER JOIN jobmaster b ON (a.fccode=b.fccode) " & _
											 "INNER JOIN jobgroupmaster c ON (Trim(b.jobgroupcode)=Trim(c.fccode)) " & _
											 "WHERE 1=1 " & _
											  filtercompany1 & _ 
											 "AND period_month=" & ptype & " AND period_year = " & tahun-1 & "" & _
											 "AND parentcode='01' "
											 
								Set oRs = Server.CreateObject("ADODB.Recordset")
								oRs.Open sqldtlupah, connstr
								if Not oRs.Eof then
									tawalkas_bi = cdbl(oRs("sawal"))
								Else
									tawalkas_bi = 0		   
								End if
								oRs.Close
								Set oRs = Nothing
								
								'cari saldo awal kas BL
								sqldtlupah = "SELECT NVL(Sum(ending_balance),0) sawal FROM " & _
											 "" & tblneraca & " a " & _
											 "INNER JOIN jobmaster b ON (a.fccode=b.fccode) " & _
											 "INNER JOIN jobgroupmaster c ON (Trim(b.jobgroupcode)=Trim(c.fccode)) " & _
											 "WHERE 1=1 " & _
											  filtercompany1 & _ 
											 "AND period_month=" & ptype & " AND period_year = " & tahun-2 & "" & _
											 "AND parentcode='01' "
											 
								Set oRs = Server.CreateObject("ADODB.Recordset")
								oRs.Open sqldtlupah, connstr
								if Not oRs.Eof then
									tawalkas_bl = cdbl(oRs("sawal"))
								Else
									tawalkas_bl = 0		   
								End if
								oRs.Close
								Set oRs = Nothing
							    
											

								'sqldtlupah = "SELECT row_id,description,sum(current_amount)current_amount,nonotes,sum(begin_amount)begin_amount,description2,rowtype " & _ 
								'			" FROM fac_laporan_management a " & _
								'			" where a.reportcode='LM_CF'" & _ 
								'			" AND period_month=" & bulan & " AND period_year=" & tahun & " " & _ 
                                '            filtercompany1 & _ 
								'			" GROUP BY row_id,description,nonotes,description2,rowtype" & _
								'			" ORDER BY row_id "
								sqldtlupah = " SELECT " & _
											 "  fchdrcode, " & _
											 "  CASE WHEN  Nvl(fcbudgetcode,'-') ='-'  THEN fchdrname " & _
											 " 	ELSE ' ' " & _
											 " 	END  fchdrname, " & _ 										
											 "  Nvl(fcbudgetcode,' ') fcbudgetcode, " & _
											 "  Nvl(parentcode,' ') parentcode," & _
											 "  Nvl(parentdesc,' ')parentdesc, " & _											
										     "  Sum(amount_bi) amount_bi, " & _
											 "  Sum(amount_bl) amount_bl, " & _
											 "  CASE WHEN  Nvl(fcbudgetcode,'-') ='-'  THEN fchdrname_eng " & _
											 "	ELSE  parentdesc_eng " & _
											 " 	END  eng " & _										
											 "  FROM " & _
											 " ( " & _
											 "  	SELECT   " & _
											 "  	fchdrcode, fchdrname,fchdrname_eng, " & _
											 " 		fcshdrcode, fcshdrname,fcshdrname_eng, " & _
											 "      parentcode, parentdesc, parentdesc_eng," & _
											 "      fcbudgetcode," & _
											 " 	    Sum(amount_bl) amount_bl, " & _
											 " 	    Sum(amount_bi) amount_bi " & _
											 " 	    FROM " & _
											 "  	( " & _
											 " 		SELECT  " & _												
											 " 		fchdrcode, fchdrname,fchdrname_eng, " & _
											 " 		fcshdrcode, fcshdrname,fcshdrname_eng, " & _
											 " 		b.fccode parentcode, b.fcname parentdesc,b.fcname_eng parentdesc_eng, " & _                    
											 " 		fchdrcode||fcshdrcode||b.fccode fcbudgetcode, " & _										 	
											 " 	    Nvl(amount_bl,0) amount_bl, " & _
											 " 		Nvl(amount_bi,0) amount_bi  " & _
											 " 		FROM " & _
											 " 		jobmaster a " & _
											 " 		INNER JOIN jobcfmaster b ON (Trim(a.job_forcashflowcode)=Trim(b.fccode)) " & _
											 " 		LEFT JOIN " & _
											 " 		( " & _
											 " 		  SELECT period_year,fccode, " & _
											 " 		  ROUND(SUM(debet)+SUM(credit),0)*-1 amount_bl,0 amount_bi " & _                                                                                    
											 " 		  FROM  " & tblneraca & "" & _
											 " 		  WHERE 1=1 " & _
											 " 		  AND period_month between 1 and " & ptype & "  AND period_year= " & tahun-1 & "  " & _
											            filtercompany1 & _ 
											 " 		  AND (debet+credit) <> 0 " & _
											 "        GROUP BY period_year,fccode " & _
											 " 		  UNION ALL " & _
											 " 		  SELECT period_year,fccode, " & _
											 " 		  0 amount_bl,ROUND(SUM(debet)+SUM(credit),0)*-1 amount_bi  " & _                                                                                    
											 " 		  FROM  " & tblneraca & " " & _
											 " 		  WHERE 1=1 " & _
											 " 		  AND period_month between 1 and " & bulan & "   AND period_year= " & tahun & "  " & _
											            filtercompany1 & _
											 " 	      AND (debet+credit) <> 0 " & _
											 "        GROUP BY period_year,fccode " & _
											 " 		)c ON (a.fccode=c.fccode) " & _
											 " 	)X " & _
											 " 	GROUP BY  fchdrcode, fchdrname,fchdrname_eng,fcshdrcode, fcshdrname,fcshdrname_eng,parentcode, parentdesc, parentdesc_eng, fcbudgetcode " & _
											 " )CATATAN " & _
											 " GROUP BY " & _
											 " GROUPING SETS ( " & _
											 " 					(fchdrcode,fchdrname,fchdrname_eng), " & _
											 " 					(fcbudgetcode,parentcode, parentdesc,parentdesc_eng,fcshdrcode,fcshdrname,fchdrcode,fchdrname) " & _
											 "   				 ) " & _
											 " ORDER BY fchdrcode,fcshdrcode,fcbudgetcode,parentcode "
'response.write "CF-" & sqldtlupah & "<br>"
								set rsdtlupah = connStr.Execute(sqldtlupah)
								no = 0
								tamountbl=0
								tamountbi=0
								While Not rsdtlupah.EOF %>
								<tr id="tru<%=no%>" >
									<td class="ratakiri"><strong><%=rsdtlupah("fchdrname")%>&nbsp;&nbsp;&nbsp;</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsdtlupah("parentdesc")%></td> <!-- head -->
									<td class="ratakanan"><%=formatnumber(rsdtlupah("amount_bi"),2)%></td>
									<td class="ratakiri"></td>
									<td class="ratakanan"><%=formatnumber(rsdtlupah("amount_bl"),2)%></td>
									<td class="ratakanan"><%if rsdtlupah("parentcode")=" " then %><strong><%=rsdtlupah("eng")%></strong><%else%><%=rsdtlupah("eng")%><%end if%></td>
									<%
									no = no +1
									if rsdtlupah("parentcode")=" " THEN
										tamountbl = tamountbl + cdbl(rsdtlupah("amount_bl"))
										tamountbi = tamountbi + cdbl(rsdtlupah("amount_bi"))										
									end if
									rsdtlupah.MoveNext
									Wend
									
									takhirkas_bi = tawalkas_bi + tamountbi
									takhirkas_bl = tawalkas_bl + tamountbl
									
									rsdtlupah.close 
									set rsdtlupah = Nothing
								%>
								</tr>
								<tr id="tru<%=no%>" >
									<td class="ratakiri"><strong>SALDO BERJALAN</strong></td> <!-- head -->
									<td class="ratakanan"><%=formatnumber(tamountbi,2)%></td>
									<td class="ratakiri"></td>
									<td class="ratakanan"><%=formatnumber(tamountbl,2)%></td>
									<td class="ratakanan"><strong>CURRENT BALANCE</strong></td>
								</tr>	
								<tr id="tru<%=no+1%>" >
									<td class="ratakiri"><strong>SALDO AWAL</strong></td> <!-- head -->
									<td class="ratakanan"><%=formatnumber(tawalkas_bi,2)%></td>
									<td class="ratakiri"></td>
									<td class="ratakanan"><%=formatnumber(tawalkas_bl,2)%></td>
									<td class="ratakanan"><strong>BEGINNING BALANCE</strong></td>
								</tr>
								<tr id="tru<%=no+2%>" >
									<td class="ratakiri"><strong>SALDO AKHIR</strong></td> <!-- head -->
									<td class="ratakanan"><%=formatnumber(takhirkas_bi,2)%></td>
									<td class="ratakiri"></td>
									<td class="ratakanan"><%=formatnumber(takhirkas_bl,2)%></td>
									<td class="ratakanan"><strong>ENDING BALANCE</strong></td>
								</tr>								
							</table>
							<br>
							<br>
							Lihat catatan atas laporan keuangan yang merupakan bagian yang tidak terpisahkan dari laporan keuangan secara keseluruhan<br>
See the accompanying notes to financial statements which are an integral part of the financial statements
						</div>
	<!-- Catatan -->
						<div class="TabbedPanelsContent">
							<table border="0" width="100%"  class='data' id='table2'>
								<tr >									
									<td style="border:none"><strong>PT. <%=CompanyName%></strong></td>
									<td style="border:none" colspan=3></td>
									<td style="border:none" align=right></td>
								</tr>
								<tr id="trheader2">									
									<td style="border:none"><strong>CATATAN ATAS LAPORAN</strong></td>
									<td style="border:none" colspan=3></td>
									<td style="border:none" align=right></td>
								</tr>
								<tr id="trheader2">									
									<td style="border:none">Per <%=textmonth(bulan)%>/<%=tahun%></td>
									<td style="border:none" colspan=3></td>
									<td style="border:none" align=right></td>
								</tr>
								<tr id="trheader2">									
									<td style="border:none">(dinyatakan dalam Rupiah)</td>
									<td style="border:none" colspan=3></td>
									<td style="border:none" align=right></td>
								</tr>
								<tr id="trheader2">									
									<td class="ratatengah">Deskripsi</td>
									<td class="ratatengah"><%=textmonth(bulan)%>/<%=tahun%></td>
									<td class="ratatengah"><%=textmonth(ptype)%>/<%=tahun-1%></td>
								</tr>
								<% 
								'get data detail upah
								'if trim(fcba) = "-" then 
								'	filtercompany1 = " AND fcba in (select fccode from businessunit where fccompanycode ='" & company & "') "
								'else 
								'	filtercompany1 = " AND fcba ='" & fcba & "' "
								'end if
								
								filtercompany1 = " AND fcba in (select fccode from businessunit b" & _ 
								" INNER JOIN USER_COMPANY_TAB u on b.FCCOMPANYCODE=u.FCCOMPANYCODE " & _ 
								" WHERE u.FCUSERCODE='" & userlogin & "' " 
								if trim(region) = "-" then 
									filtercompany1 = filtercompany1 & " AND b.region like '%' "
								else 
									filtercompany1 = filtercompany1 & " AND b.region = '" & region & "' "
								end if

								if trim(company) = "-" then 
									filtercompany1 = filtercompany1 & " AND b.fccompanycode like '%' "
								else 
									filtercompany1 = filtercompany1 & " AND b.fccompanycode = '" & company & "' "
								end if

								if trim(fcba) = "-" then 
									filtercompany1 = filtercompany1 & " AND b.fccode like '%' )"
								else 
									filtercompany1 = filtercompany1 & " AND b.fccode ='" & fcba & "' )"
								end if
								
                               sqldtlupah = "	SELECT " & _
											"	parentcode," & _
											"	CASE " & _
											"	WHEN  Nvl(groupcode,'-')='-' AND  Nvl(fccode,'-') ='-'   THEN parentcode||'-'||parentdesc " & _
											"	ELSE ' ' " & _
											"	END  parentdesc," & _ 
											"	Nvl(groupcode,' ') groupcode," & _
											"	CASE " & _
											"	WHEN  Nvl(fccode,'-') ='-'   THEN Nvl(groupname,' ') " & _
											"	ELSE ' ' " & _
											"	END  groupname," & _
											"	Nvl(fccode,' ')fccode,Nvl(fcname,' ')fcname," & _
											"	Sum(begin_amount) begin_amount," & _
											"	Sum(current_amount) current_amount " & _
											"	FROM " & _
											"	(" & _
											"	SELECT parentcode,parentdesc,groupcode,groupname,fccode,fcname,Sum(begin_amount) begin_amount,Sum(current_amount)current_amount FROM " & _
											"	( " & _
											"	SELECT " & _
											"	b.parentcode,b.parentdesc," & _
											"	b.parentcode||b.fccode groupcode,b.fcname groupname ," & _
											"	c.fccode ,a.fcname,c.ending_balance begin_amount,0 current_amount " & _
											"	FROM " & _
											"		jobmaster a" & _
											"		LEFT JOIN jobgroupmaster b ON (a.jobgroupcode=b.fccode)" & _
											"		LEFT JOIN (SELECT period_year,period_month,fcba,fccode,CASE WHEN SubStr(fccode,1,1) = '1' then ending_balance  ELSE ending_balance*-1 END ending_balance " & _ 
											"       FROM " & tblneraca & " WHERE period_month=" & ptype & " AND period_year=" & tahun-1 & ") c ON (a.fccode=c.fccode)" & _
											"		WHERE ending_balance <> 0 " & _
													filtercompany1  & _
											"		AND period_month=" & ptype & "" & _
											"		AND period_year=" & tahun-1 & " " & _
											"	UNION ALL " & _
											"	SELECT " & _
											"	b.parentcode,b.parentdesc," & _
											"	b.parentcode||b.fccode groupcode,b.fcname groupname ," & _
											"	c.fccode ,a.fcname,0 begin_amount,c.ending_balance current_amount " & _
											"	FROM " & _
											"		jobmaster a " & _
											"		LEFT JOIN jobgroupmaster b ON (a.jobgroupcode=b.fccode)" & _
											"		LEFT JOIN (SELECT period_year,period_month,fcba,fccode,CASE WHEN SubStr(fccode,1,1) = '1' then ending_balance  ELSE ending_balance*-1 END ending_balance " & _ 
											"       FROM " & tblneraca & " WHERE period_month=" & bulan & " AND period_year=" & tahun & ") c ON (a.fccode=c.fccode)" & _
											"		WHERE ending_balance <> 0 " & _
													filtercompany1  & _
											"		AND period_month=" & bulan & "" & _
											"		AND period_year=" & tahun & ""  & _
											"	)X " & _
											"	GROUP BY  parentcode,parentdesc,groupcode,groupname,fccode,fcname " & _
											"	)CATATAN " & _
											"	GROUP BY " & _
											"	  GROUPING SETS (" & _
											"		(parentcode, parentdesc)," & _
											"		(groupcode, groupname, parentcode, parentdesc)," & _
											"		(fccode, fcname, groupcode, groupname, parentcode, parentdesc)" & _
											"	  )" & _
											"	ORDER BY parentcode,groupcode,fccode"
								'response.write "X-" &  sqldtlupah
								set rsdtlupah = connStr.Execute(sqldtlupah)
								no = 0
								While Not rsdtlupah.EOF %>
								<% if rsdtlupah("parentdesc")<>" " then %>
								<tr id="tru<%=no%>" class="top bottom row">
								<%else%>
								<tr id="tru<%=no%>" >
								<%end if%>
								
									<td class="ratakiri"><strong><%=rsdtlupah("parentdesc")%></strong>&nbsp;&nbsp;&nbsp;<%=rsdtlupah("groupname")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsdtlupah("fccode")%>&nbsp;<%=rsdtlupah("fcname")%></td> <!-- head -->
									<td class="ratakanan"><%=formatnumber(rsdtlupah("current_amount"),2)%></td> <!-- Current -->
									<td class="ratakanan"><%=formatnumber(rsdtlupah("begin_amount"),2)%></td> <!-- begin -->
									
									<%
									no = no +1
									rsdtlupah.MoveNext
									Wend
									rsdtlupah.close 
									set rsdtlupah = Nothing
								%>
								</tr>
							</table>
							<br>
							<br>
							
						</div>
						
						
<!-- Lap produksi -->
						<div class="TabbedPanelsContent">
							<table border="0" width="100%"  class='data' id='table2'>
								<tr >									
									<td style="border:none"><strong>PT. <%=CompanyName%></strong></td>
									<td style="border:none" colspan=3></td>
								</tr>
								<tr id="trheader2">									
									<td style="border:none"><strong>LAPORAN PRODUKSI</strong></td>
									<td style="border:none" colspan=3></td>
								</tr>
								<tr id="trheader2">									
									<td style="border:none">Per <%=textmonth(bulan)%>/<%=tahun%></td>
									<td style="border:none" colspan=3></td>
								</tr>
								<tr id="trheader2">									
									<td class="ratatengah">Deskripsi</td> <!-- head -->
									<td class="ratatengah"><%=textmonth(bulanbfr)%>/<%=tahunbfr%></td> <!-- Sub head -->
									<td class="ratatengah"><%=textmonth(bulan)%>/<%=tahun%></td> <!-- Detail -->
									<td class="ratatengah">s/d <%=textmonth(bulan)%>/<%=tahun%></td> <!-- Current -->
								</tr>
								<% 
								'if trim(fcba) = "-" then 
								'	filtercompany1 = " AND fcba in (select fccode from businessunit where fccompanycode ='" & company & "') "
								'else 
								'	filtercompany1 = " AND fcba ='" & fcba & "' "
								'end if
								
								filtercompany1 = " AND fcba in (select fccode from businessunit b" & _ 
								" INNER JOIN USER_COMPANY_TAB u on b.FCCOMPANYCODE=u.FCCOMPANYCODE " & _ 
								" WHERE u.FCUSERCODE='" & userlogin & "' " 
								if trim(region) = "-" then 
									filtercompany1 = filtercompany1 & " AND b.region like '%' "
								else 
									filtercompany1 = filtercompany1 & " AND b.region = '" & region & "' "
								end if

								if trim(company) = "-" then 
									filtercompany1 = filtercompany1 & " AND b.fccompanycode like '%' "
								else 
									filtercompany1 = filtercompany1 & " AND b.fccompanycode = '" & company & "' "
								end if

								if trim(fcba) = "-" then 
									filtercompany1 = filtercompany1 & " AND b.fccode like '%' )"
								else 
									filtercompany1 = filtercompany1 & " AND b.fccode ='" & fcba & "' )"
								end if

								sqldtlupah = "SELECT row_id,description,case when row_id > 29 then '' else acccode end acccode ,sum(nvl(current_amount,0))current_amount,sum(nvl(begin_amount,0))begin_amount,sum(nvl(until_amount,0))until_amount,rowtype " & _ 
											" FROM fac_laporan_management a " & _
											" where a.reportcode='LM_PROD'" & _ 
											" AND period_month=" & bulan & " AND period_year=" & tahun & " " & _ 
                                            filtercompany1 & _ 
											" GROUP BY row_id,description,acccode,rowtype" & _
											" ORDER BY row_id "

								set rsdtlupah = connStr.Execute(sqldtlupah)
								no = 0
								While Not rsdtlupah.EOF %>
								<% if rsdtlupah("rowtype")="3" then %>
								<tr id="tru<%=no%>" class="top bottom row">
								<%else%>
								<tr id="tru<%=no%>" >
								<%end if%>
								 
									
									<td class="ratakiri"><%if rsdtlupah("rowtype")="2" then %><strong><%=rsdtlupah("description")%></strong><%else%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsdtlupah("acccode")%>&nbsp;&nbsp;&nbsp;<%=rsdtlupah("description")%><%end if%></td>
									<td class="ratakanan"><%if rsdtlupah("rowtype")="2" then %>&nbsp;<%else%><%=formatnumber(rsdtlupah("begin_amount"),2)%><%end if%></td>
									<td class="ratakanan"><%if rsdtlupah("rowtype")="2" then %>&nbsp;<%else%><%=formatnumber(rsdtlupah("current_amount"),2)%><%end if%></td>
									<td class="ratakanan"><%if rsdtlupah("rowtype")="2" then %>&nbsp;<%else%><%=formatnumber(rsdtlupah("until_amount"),2)%><%end if%></td>
								
								<%
									no = no +1
									rsdtlupah.MoveNext
									Wend
									rsdtlupah.close 
									set rsdtlupah = Nothing
								%>
								
							</table>
								<br>
							<br>
							
						</div>
					
					

<!-- Realisasi Budget -->
						<div class="TabbedPanelsContent">
<div id="wrapper">

							<table border="0" width="100%"  class="xthf-site"  id="">
							<thead>
								<tr id="trheader2">									
									<th colspan=22 style=" background-color:#FFFFFF; color:#000000; border:none">
									<div id="header-wrapper">
										<div id="head_judul" style="float:left; text-align:left;">
										<h4>PT. <%=CompanyName%></h4>
										</div>
										<div id="subheader"  style="width:50%; margin:auto;">
										<h4 class="judul">Realisasi & Budget</h1>
										<h4>Per <%=textmonth(bulan)%>/<%=tahun%></h5>
										</div>
										<div style="clear:both;"></div>
									</div>
									</th>
								</tr>
								
								<tr style="background-color:#00A300;color:#fff;">
									<th rowspan=4>Deskripsi</th>
								</tr>
								<tr style="background-color:#00A300;color:#fff;">
									<th colspan="7">Bulan Lalu</th>
									<th colspan="7">Bulan Ini</th>
									<th colspan="7">s/d Bulan Ini</th>
								</tr>
								<tr style="background-color:#00A300;color:#fff;">
									<th colspan="3">Realisasi</th>
									<th colspan="3">RKB</th>
									<th >Var</th>
									
									<th colspan="3">Realisasi</th>
									<th colspan="3">RKB</th>
									<th >Var</th>
									
									<th colspan="3">Realisasi</th>
									<th colspan="3">RKB</th>	
									<th >Var</th>
																		
								</tr>
								<tr style="background-color:#00A300;color:#fff;">	
                                    <th>Qty</th> <!-- BL.Rel -->
									<th>Rp</th> 
									<th>Rp/Qty</th> 
									<th>Qty</th> <!-- BL.RKB -->
									<th>Rp</th> 
									<th>Rp/Qty</th> 
									<th>%</th> 
									
									
									<th>Qty</th> <!-- BL.Rel -->
									<th>Rp</th> 
									<th>Rp/Qty</th> 
									<th>Qty</th> <!-- BL.RKB -->
									<th>Rp</th> 
									<th>Rp/Qty</th> 
									<th>%</th> 
									
									<th>Qty</th> <!-- BL.Rel -->
									<th>Rp</th> 
									<th>Rp/Qty</th> 
									<th>Qty</th> <!-- BL.RKB -->
									<th>Rp</th> 
									<th>Rp/Qty</th> 
									<th>%</th> 
									
								</tr>
								</thead>
								
								 <% 
								'get data detail upah
								'if trim(fcba) = "-" then 
								'	filtercompany1 = " AND fcba in (select fccode from businessunit where fccompanycode ='" & company & "') "
								'else 
								'	filtercompany1 = " AND fcba ='" & fcba & "' "
								'end if
								
								filtercompany1 = " AND fcba in (select fccode from businessunit b" & _ 
								" INNER JOIN USER_COMPANY_TAB u on b.FCCOMPANYCODE=u.FCCOMPANYCODE " & _ 
								" WHERE u.FCUSERCODE='" & userlogin & "' " 
								if trim(region) = "-" then 
									filtercompany1 = filtercompany1 & " AND b.region like '%' "
								else 
									filtercompany1 = filtercompany1 & " AND b.region = '" & region & "' "
								end if

								if trim(company) = "-" then 
									filtercompany1 = filtercompany1 & " AND b.fccompanycode like '%' "
								else 
									filtercompany1 = filtercompany1 & " AND b.fccompanycode = '" & company & "' "
								end if

								if trim(fcba) = "-" then 
									filtercompany1 = filtercompany1 & " AND b.fccode like '%' )"
								else 
									filtercompany1 = filtercompany1 & " AND b.fccode ='" & fcba & "' )"
								end if
								
                               sqldtlupah = " SELECT " & _
											" fchdrcode, " & _
											" CASE WHEN  Nvl(fcsbhdrcode,'-')='-' AND  Nvl(fcssbhdrcode,'-') ='-' AND  Nvl(fcbudgetcode,'-') ='-'  THEN fchdrname " & _
											" 						ELSE ' ' " & _
											" 						END  fchdrname, " & _ 										
											" Nvl(fcsbhdrcode,' ')  fcsbhdrcode, " & _
											" CASE WHEN   Nvl(fcssbhdrcode,'-') ='-' AND  Nvl(fcbudgetcode,'-') ='-'  THEN Nvl(fcsbhdrname,' ') " & _
											" 						ELSE ' ' " & _
											" 						END  fcsbhdrname, " & _ 										
											" Nvl(fcssbhdrcode,' ') fcssbhdrcode, " & _
											" CASE WHEN  Nvl(fcbudgetcode,'-') ='-'  THEN Nvl(fcssbhdrname,' ') " & _
											" 						ELSE ' ' " & _
											" 						END  fcssbhdrname, " & _
											" Nvl(fcbudgetcode,' ') fcbudgetcode,CASE WHEN Nvl(jobcode,'-') = '-' THEN Nvl(fcname,' ') ELSE ' ' end fcname, " & _
											" Nvl(jobcode,' ') jobcode,Nvl(jobname,' ') jobname, " & _
										    " " & _

											" CASE WHEN MAX(period_month) = 1 THEN 0 ELSE SUM(output_bl) END AS output_bl, " & _
											" CASE WHEN MAX(period_month) = 1 THEN 0 ELSE ROUND(SUM(amount_bl),0) END AS amount_bl, " & _
											" CASE WHEN MAX(period_month) = 1 THEN 0 ELSE Round(Decode(Sum(output_bl),0,0,Sum(amount_bl)/Sum(output_bl)),0) END rpqty_bl, " & _
											" CASE WHEN MAX(period_month) = 1 THEN 0 ELSE Sum(rkb_output_bl) END rkb_output_bl, " & _
											" CASE WHEN MAX(period_month) = 1 THEN 0 ELSE Round(Sum(rkb_amount_bl),0) END rkb_amount_bl, " & _
											" CASE WHEN MAX(period_month) = 1 THEN 0 ELSE Round(Decode(Sum(rkb_output_bl),0,0,Sum(rkb_amount_bl)/Sum(rkb_output_bl)),0) END rkb_rpqty_bl, " & _
											" CASE WHEN MAX(period_month) = 1 THEN 0 ELSE  " & _
											" Decode(Round(Decode(Sum(rkb_output_bl),0,0,Sum(rkb_amount_bl)/Sum(rkb_output_bl)),0),0,0,(Round(Decode(Sum(output_bl),0,0,Sum(amount_bl)/Sum(output_bl)),0)/Round(Decode(Sum(rkb_output_bl),0,0, " & _
											" Sum(rkb_amount_bl)/Sum(rkb_output_bl)),0))*100) END var_bl,  " & _

											"    Sum(output_bi) output_bi,Round(Sum(amount_bi),0) amount_bi,Round(Decode(Sum(output_bi),0,0,Sum(amount_bi)/Sum(output_bi)),0) rpqty_bi, " & _
											"    Sum(rkb_output_bi) rkb_output_bi,Round(Sum(rkb_amount_bi),0) rkb_amount_bi,Round(Decode(Sum(rkb_output_bi),0,0,Sum(rkb_amount_bi)/Sum(rkb_output_bi)),0) rkb_rpqty_bi, " & _
											"    Decode(Round(Decode(Sum(rkb_output_bi),0,0,Sum(rkb_amount_bi)/Sum(rkb_output_bi)),0),0,0,(Round(Decode(Sum(output_bi),0,0,Sum(amount_bi)/Sum(output_bi)),0)/Round(Decode(Sum(rkb_output_bi),0,0,Sum(rkb_amount_bi)/Sum(rkb_output_bi)),0))*100) var_bi," & _
											"    " & _
											"    Sum(output_sbi) output_sbi,Round(Sum(amount_sbi),0) amount_sbi,Round(Decode(Sum(output_sbi),0,0,Sum(amount_sbi)/Sum(output_sbi)),0) rpqty_sbi, " & _
											"    Sum(rkb_output_sbi) rkb_output_sbi,Round(Sum(rkb_amount_sbi),0) rkb_amount_sbi,Round(Decode(Sum(rkb_output_sbi),0,0,Sum(rkb_amount_sbi)/Sum(rkb_output_sbi)),0) rkb_rpqty_sbi, " & _
											"    Decode(Round(Decode(Sum(rkb_output_sbi),0,0,Sum(rkb_amount_sbi)/Sum(rkb_output_sbi)),0),0,0,(Round(Decode(Sum(output_sbi),0,0,Sum(amount_sbi)/Sum(output_sbi)),0)/Round(Decode(Sum(rkb_output_sbi),0,0,Sum(rkb_amount_sbi)/Sum(rkb_output_sbi)),0))*100) var_sbi" & _
											
											" FROM " & _
											" ( " & _
											"  	SELECT " & _  
											"  	fchdrcode, fchdrname,fchdrname_eng, " & _
											"  	fcsbhdrcode, fcsbhdrname,fcsbhdrname_eng, " & _                      
											"  	fcssbhdrcode,fcssbhdrname,fcssbhdrname_eng, " & _
											"  	fcbudgetcode,fcname,jobcode,jobname,period_month," & _
											" " & _
											" 	Sum(output_bl) output_bl,Sum(amount_bl) amount_bl,Sum(rkb_output_bl) rkb_output_bl,Sum(rkb_amount_bl) rkb_amount_bl, " & _
											" 	Sum(output_bi) output_bi,Sum(amount_bi) amount_bi,Sum(rkb_output_bi) rkb_output_bi,Sum(rkb_amount_bi) rkb_amount_bi, " & _
											" 	Sum(output_sbi) output_sbi,Sum(amount_sbi) amount_sbi,Sum(rkb_output_sbi) rkb_output_sbi,Sum(rkb_amount_sbi) rkb_amount_sbi " & _
											" " & _
											" 	FROM " & _
											"  	( " & _
											" 		SELECT  " & _												
											" 		fchdrcode, fchdrname,fchdrname_eng, " & _
											" 		fcsbhdrcode, fcsbhdrname,fcsbhdrname_eng, " & _                      
											" 		fcssbhdrcode,fcssbhdrname,fcssbhdrname_eng, " & _
											" 		fchdrcode||fcsbhdrcode||fcssbhdrcode||b.fccode fcbudgetcode, " & _
											" 		b.fcname,  " & _
											" 		a.fccode jobcode,a.fcname jobname,period_month, " & _
											"  " & _
											" 		Nvl(output_bl,0) output_bl,Nvl(amount_bl,0) amount_bl,Nvl(rkb_output_bl,0) rkb_output_bl,Nvl(rkb_amount_bl,0) rkb_amount_bl, " & _
											" 		Nvl(output_bi,0) output_bi,Nvl(amount_bi,0) amount_bi,Nvl(rkb_output_bi,0) rkb_output_bi,Nvl(rkb_amount_bi,0) rkb_amount_bi, " & _
											" 		Nvl(output_sbi,0) output_sbi,Nvl(amount_sbi,0) amount_sbi,Nvl(rkb_output_sbi,0) rkb_output_sbi,Nvl(rkb_amount_sbi,0) rkb_amount_sbi " & _ 
											" 		FROM " & _
											" 		jobmaster a " & _
											" 		INNER JOIN jobbtmaster b ON (Trim(a.job_forbudgetcode)=Trim(b.fccode)) " & _
											" 		LEFT JOIN " & _
											" 		( " & _
											" 		  SELECT period_year,period_month,fcba,fccode, " & _
											" 		  0 output_bl, Sum(debet)+sum(credit) amount_bl,0 rkb_output_bl,0 rkb_amount_bl, " & _
											" 		  0 output_bi,0 amount_bi,0 rkb_output_bi,0 rkb_amount_bi, " & _
											" 		  0 output_sbi,0 amount_sbi,0 rkb_output_sbi,0 rkb_amount_sbi  " & _                                                                                   
											" 		  FROM  " & tblneraca & " " & _
											" 		  WHERE 1=1  AND fccode in (SELECT fccode FROM jobmaster WHERE job_forbudgetcode IN (SELECT fccode FROM jobbtmaster  WHERE fchdrcode='01')) " & _ 
											"		  AND period_month=" & bulanbfr & " AND period_year=" & tahunbfr & "" & _ 
											filtercompany1   & _
											" 		  GROUP BY period_year,period_month,fcba,fccode " & _
											" 		  UNION ALL " & _
											" 		  SELECT period_year,period_month,fcba,fccode, " & _
											" 		  0 output_bl, 0 amount_bl,0 rkb_output_bl,0 rkb_amount_bl, " & _
											" 		  0 output_bi,Sum(debet)+sum(credit) amount_bi,0 rkb_output_bi,0 rkb_amount_bi, " & _
											" 		  0 output_sbi,0 amount_sbi,0 rkb_output_sbi,0 rkb_amount_sbi " & _                                                                                    
											" 		  FROM  " & tblneraca & " " & _
											" 		  WHERE 1=1 AND fccode in (SELECT fccode FROM jobmaster WHERE job_forbudgetcode IN (SELECT fccode FROM jobbtmaster  WHERE fchdrcode='01')) " & _ 
											" 		 AND period_month=" & bulan & "  AND period_year=" & tahun & "" & _ 
											filtercompany1   & _
											" 		  GROUP BY period_year,period_month,fcba,fccode " & _
											" 		  UNION ALL " & _
											" 		  SELECT period_year,period_month,fcba,fccode, " & _
											" 		  0 output_bl, 0 amount_bl,0 rkb_output_bl,0 rkb_amount_bl, " & _
											" 		  0 output_bi, 0 amount_bi,0 rkb_output_bi,0 rkb_amount_bi, " & _
											" 		  0 output_sbi,Sum(debet)+sum(credit) amount_sbi,0 rkb_output_sbi,0 rkb_amount_sbi " & _
											" 		  FROM  " & tblneraca & " " & _
											" 		  WHERE 1=1 AND fccode in (SELECT fccode FROM jobmaster WHERE job_forbudgetcode IN (SELECT fccode FROM jobbtmaster  WHERE fchdrcode='01')) " & _ 
											" 		  AND period_month between 1 AND " & bulan & " AND period_year=" & tahun & "" & _ 
											filtercompany1  & _
											" 		  GROUP BY period_year,period_month,fcba,fccode " & _
											" 		  UNION ALL " & _
											" 		  SELECT period_year,period_month,fcba,fccode, " & _
											" 		  0 output_bl, Sum(debet)+sum(credit) amount_bl,0 rkb_output_bl,0 rkb_amount_bl, " & _
											" 		  0 output_bi,0 amount_bi,0 rkb_output_bi,0 rkb_amount_bi, " & _
											" 		  0 output_sbi,0 amount_sbi,0 rkb_output_sbi,0 rkb_amount_sbi  " & _                                                                                   
											" 		  FROM  " & tblneraca & " " & _
											" 		  WHERE 1=1 AND fccode in (SELECT fccode FROM jobmaster WHERE job_forbudgetcode IN (SELECT fccode FROM jobbtmaster  WHERE fchdrcode='02')) " & _
											"		  AND period_month=" & bulanbfr & " AND period_year=" & tahunbfr & "" & _ 
											filtercompany1   & _
											" 		  GROUP BY period_year,period_month,fcba,fccode " & _
											" 		  UNION ALL " & _
											" 		  SELECT period_year,period_month,fcba,fccode, " & _
											" 		  0 output_bl, 0 amount_bl,0 rkb_output_bl,0 rkb_amount_bl, " & _
											" 		  0 output_bi,Sum(debet)+sum(credit) amount_bi,0 rkb_output_bi,0 rkb_amount_bi, " & _
											" 		  0 output_sbi,0 amount_sbi,0 rkb_output_sbi,0 rkb_amount_sbi " & _                                                                                    
											" 		  FROM  " & tblneraca & " " & _
											" 		  WHERE 1=1  AND fccode in (SELECT fccode FROM jobmaster WHERE job_forbudgetcode IN (SELECT fccode FROM jobbtmaster  WHERE fchdrcode='02')) " & _ 
											" 		  AND period_month=" & bulan & "  AND period_year=" & tahun & "" & _ 
											filtercompany1   & _
											" 		  GROUP BY period_year,period_month,fcba,fccode " & _
											" 		  UNION ALL " & _
											" 		  SELECT period_year,period_month,fcba,fccode, " & _
											" 		  0 output_bl, 0 amount_bl,0 rkb_output_bl,0 rkb_amount_bl, " & _
											" 		  0 output_bi, 0 amount_bi,0 rkb_output_bi,0 rkb_amount_bi, " & _
											" 		  0 output_sbi,Sum(debet)+sum(credit) amount_sbi,0 rkb_output_sbi,0 rkb_amount_sbi " & _
											" 		  FROM  " & tblneraca & " " & _
											" 		  WHERE 1=1  AND fccode in (SELECT fccode FROM jobmaster WHERE job_forbudgetcode IN (SELECT fccode FROM jobbtmaster  WHERE fchdrcode='02')) " & _ 
											"		  AND period_month between 1 AND " & bulan & " AND period_year=" & tahun & "" & _ 
											filtercompany1  & _
											" 		  GROUP BY period_year,period_month,fcba,fccode " & _
											"		  UNION ALL " & _
											"		  SELECT period_year,period_month,fcba,Trim(acccode) fccode, Sum(begin_amount) output_bl, " & _
											"		  0 amount_bl,0 rkb_output_bl,0 rkb_amount_bl, " & _
											"		  Sum(current_amount) output_bi, 0 amount_bi,0 rkb_output_bi,0 rkb_amount_bi, " & _
											"		  Sum(until_amount) output_sbi,0 amount_sbi,0 rkb_output_sbi,0 rkb_amount_sbi " & _
											"		  FROM  fac_laporan_management  " & _
											"		  WHERE reportcode='LM_PROD'  AND acccode <> '-'  " & _
											"		  AND trim(acccode) in (SELECT fccode FROM jobmaster WHERE job_forbudgetcode IN (SELECT fccode FROM jobbtmaster)) " & _
											"		  AND period_month = " & bulan & "  AND period_year= " & tahun & "" & _
											filtercompany1  & _
											"		  GROUP BY  period_year,period_month,fcba,acccode " & _ 
											"		  UNION ALL " & _
											"		  SELECT period_year,period_month,a.fcba,trim(actcode) fccode, " & _ 
											"		  0 output_bl, 0 amount_bl,Sum(volume) rkb_output_bl, Sum(totalamount_hk) rkb_amount_bl, " & _
											"		  0 output_bi, 0 amount_bi, 0 rkb_output_bi, 0 rkb_amount_bi, " & _
											"		  0 output_sbi,0 amount_sbi,0 rkb_output_sbi,0 rkb_amount_sbi " & _
											"		  FROM rkbhdr a INNER JOIN rkbupahdtl b ON (a.rkbno=b.rkbno AND a.fcba=b.fcba) " & _ 
											"		  WHERE " & _ 
											"		  trim(actcode) in (SELECT fccode FROM jobmaster WHERE job_forbudgetcode IN (SELECT fccode FROM jobbtmaster)) " & _
											"		  AND period_month = " & bulanbfr & "  AND period_year= " & tahunbfr & "" & _
											"		  GROUP BY  period_year,period_month,a.fcba,actcode  " & _ 
											"		  UNION ALL " & _
											"		  SELECT period_year,period_month,a.fcba,trim(actcode) fccode, " & _ 
											"		  0 output_bl, 0 amount_bl,0 rkb_output_bl,0 rkb_amount_bl, " & _
											"		  0 output_bi, 0 amount_bi, Sum(volume) rkb_output_bi,Sum(totalamount_hk)  rkb_amount_bi, " & _
											"		  0 output_sbi,0 amount_sbi,0 rkb_output_sbi,0 rkb_amount_sbi " & _
											"		  FROM rkbhdr a INNER JOIN rkbupahdtl b ON (a.rkbno=b.rkbno AND a.fcba=b.fcba) " & _ 
											"		  WHERE " & _ 
											"		  trim(actcode) in (SELECT fccode FROM jobmaster WHERE job_forbudgetcode IN (SELECT fccode FROM jobbtmaster)) " & _
											"		  AND period_month = " & bulan & "  AND period_year= " & tahun & "" & _
											"		  GROUP BY  period_year,period_month,a.fcba,actcode  " & _ 
											"		  UNION ALL " & _
											"		  SELECT period_year,period_month,a.fcba,trim(actcode) fccode, " & _ 
											"		  0 output_bl, 0 amount_bl,0 rkb_output_bl,0 rkb_amount_bl, " & _
											"		  0 output_bi, 0 amount_bi, 0 rkb_output_bi,0 rkb_amount_bi, " & _
											"		  0 output_sbi,0 amount_sbi,Sum(volume) rkb_output_sbi,Sum(totalamount_hk)  rkb_amount_sbi " & _
											"		  FROM rkbhdr a INNER JOIN rkbupahdtl b ON (a.rkbno=b.rkbno AND a.fcba=b.fcba) " & _ 
											"		  WHERE " & _ 
											"		  trim(actcode) in (SELECT fccode FROM jobmaster WHERE job_forbudgetcode IN (SELECT fccode FROM jobbtmaster)) " & _
											"		  AND period_month between 1 AND " & bulan & "  AND period_year= " & tahun & "" & _
											"		  GROUP BY  period_year,period_month,a.fcba,actcode  " & _ 
											" 		)c ON (a.fccode=c.fccode) " & _
											" 	)X " & _
											" 	GROUP BY  fchdrcode,fchdrname,fchdrname_eng,fcsbhdrcode,fcsbhdrname,fcsbhdrname_eng,fcssbhdrcode,fcssbhdrname,fcssbhdrname_eng,fcbudgetcode,fcname,jobcode,jobname,period_month " & _
											" )CATATAN " & _
											" GROUP BY " & _
											" GROUPING SETS ( " & _
											" 					(fchdrcode,fchdrname), " & _
											" 					(fcsbhdrcode,fcsbhdrname,fchdrcode,fchdrname), " & _
											" 					(fcssbhdrcode,fcssbhdrname,fcsbhdrcode,fcsbhdrname,fchdrcode,fchdrname), " & _
											"  					(fcbudgetcode,fcname,fcssbhdrcode,fcssbhdrname,fcsbhdrcode,fcsbhdrname,fchdrcode,fchdrname), " & _
											" 					(jobcode,jobname,fcbudgetcode,fcname,fcssbhdrcode,fcssbhdrname,fcsbhdrcode,fcsbhdrname,fchdrcode,fchdrname) " & _
											"   				 ) having Sum(amount_sbi) + Sum(rkb_amount_sbi) <> 0 " & _
											" ORDER BY fchdrcode,fcsbhdrcode,fcssbhdrcode,fcbudgetcode,jobcode"
							''	response.write sqldtlupah 'tes commit
								set rsdtlupah = connStr.Execute(sqldtlupah)
								
								no = 0
								tamountbl=0
								tamountbi=0
								tamountsbi=0
								While Not rsdtlupah.EOF 
								%>
								<tr class='' >
									<td class="ratakiri"  style="white-space: nowrap;"><strong><%=rsdtlupah("fchdrname")%><%=rsdtlupah("fcsbhdrname")%><%=rsdtlupah("fcssbhdrname")%><%=rsdtlupah("fcname")%></strong><%=rsdtlupah("jobcode")%>&nbsp;<%=rsdtlupah("jobname")%></td> <!-- head -->
									<!-- BL -->
									<td class="ratakanan"><%=formatnumber(rsdtlupah("output_bl"),2)%></td> <!-- Qty.Rel -->
									<td class="ratakanan"><%=formatnumber(rsdtlupah("amount_bl"),2)%></td> <!-- Rp.Rel -->
									<td class="ratakanan"><%=formatnumber(rsdtlupah("rpqty_bl"),2)%></td> <!-- Rp/Qty.Rel -->
									
									<td class="ratakanan"><%=formatnumber(rsdtlupah("rkb_output_bl"),2)%></td> <!-- Qty.Bud -->
									<td class="ratakanan"><%=formatnumber(rsdtlupah("rkb_amount_bl"),2)%></td> <!-- Rp.Bud -->
									<td class="ratakanan"><%=formatnumber(rsdtlupah("rkb_rpqty_bl"),2)%></td> <!-- Rp/Qty.Bud -->
									
									<td class="ratakanan"><%=formatnumber(rsdtlupah("var_bl"),2)%></td> <!-- BL.Var -->
									
									<!-- BI -->
									<td class="ratakanan"><%=formatnumber(rsdtlupah("output_bi"),2)%></td> <!-- Qty.Rel -->
									<td class="ratakanan"><%=formatnumber(rsdtlupah("amount_bi"),2)%></td> <!-- Rp.Rel -->
									<td class="ratakanan"><%=formatnumber(rsdtlupah("rpqty_bi"),2)%></td> <!-- Rp/Qty.Rel -->
									
									<td class="ratakanan"><%=formatnumber(rsdtlupah("rkb_output_bi"),2)%></td> <!-- Qty.Bud -->
									<td class="ratakanan"><%=formatnumber(rsdtlupah("rkb_amount_bi"),2)%></td> <!-- Rp.Bud -->
									<td class="ratakanan"><%=formatnumber(rsdtlupah("rkb_rpqty_bi"),2)%></td> <!-- Rp/Qty.Bud -->
									
									
									<td class="ratakanan"><%=formatnumber(rsdtlupah("var_bi"),2)%></td> <!-- BI.Var -->
									
									<!-- SBI -->
									<td class="ratakanan"><%=formatnumber(rsdtlupah("output_sbi"),2)%></td> <!-- Qty.Rel -->
									<td class="ratakanan"><%=formatnumber(rsdtlupah("amount_sbi"),2)%></td> <!-- Rp.Rel -->
									<td class="ratakanan"><%=formatnumber(rsdtlupah("rpqty_sbi"),2)%></td> <!-- Rp/Qty.Rel -->
									
									<td class="ratakanan"><%=formatnumber(rsdtlupah("rkb_output_sbi"),2)%></td> <!-- Qty.Bud -->
									<td class="ratakanan"><%=formatnumber(rsdtlupah("rkb_amount_sbi"),2)%></td> <!-- Rp.Bud -->
									<td class="ratakanan"><%=formatnumber(rsdtlupah("rkb_rpqty_sbi"),2)%></td> <!-- Rp/Qty.Bud -->
									
									
									<td class="ratakanan"><%=formatnumber(rsdtlupah("var_sbi"),2)%></td> <!-- SBI.Var -->
									
									
									
									<%
									no = no +1
									if rsdtlupah("fcsbhdrcode")=" " and rsdtlupah("jobcode")=" " THEN
										tamountbl = tamountbl - cdbl(rsdtlupah("amount_bl"))
										tamountbi = tamountbi - cdbl(rsdtlupah("amount_bi"))
										tamountsbi = tamountsbi - cdbl(rsdtlupah("amount_sbi"))
									end if
									rsdtlupah.MoveNext
									Wend
									rsdtlupah.close 
									set rsdtlupah = Nothing
								%>
								</tr>
								<tr class="huruftebal ">
									<td class="ratakiri" >BALANCE</td> <!-- head -->
									<!-- BL -->
									<td class="ratakanan"></td> <!-- Qty.Rel -->
									<td class="ratakanan"><%=formatnumber(tamountbl,2)%></td> <!-- Rp.Rel -->
									<td class="ratakanan"></td> <!-- Rp/Qty.Rel -->
									
									<td class="ratakanan"></td> <!-- Qty.Bud -->
									<td class="ratakanan"></td> <!-- Rp.Bud -->
									<td class="ratakanan"></td> <!-- Rp/Qty.Bud -->
									
									<td class="ratakanan"></td> <!-- BL.Var -->
									
									<!-- BI -->
									<td class="ratakanan"></td> <!-- Qty.Rel -->
									<td class="ratakanan"><%=formatnumber(tamountbi,2)%></td> <!-- Rp.Rel -->
									<td class="ratakanan"></td> <!-- Rp/Qty.Rel -->
									
									<td class="ratakanan"></td> <!-- Qty.Bud -->
									<td class="ratakanan"></td> <!-- Rp.Bud -->
									<td class="ratakanan"></td> <!-- Rp/Qty.Bud -->
									
									
									<td class="ratakanan"></td> <!-- BI.Var -->
									
									<!-- SBI -->
									<td class="ratakanan"></td> <!-- Qty.Rel -->
									<td class="ratakanan"><%=formatnumber(tamountsbi,2)%></td> <!-- Rp.Rel -->
									<td class="ratakanan"></td> <!-- Rp/Qty.Rel -->
									
									<td class="ratakanan"></td> <!-- Qty.Bud -->
									<td class="ratakanan"></td> <!-- Rp.Bud -->
									<td class="ratakanan"></td> <!-- Rp/Qty.Bud -->
									
									
									<td class="ratakanan"></td> <!-- SBI.Var -->
								</tr>
								
							</table>
								<br>
							<br>
</div>
						</div>
					</div>
					
					
				</div>
			</td>
	</tr>
			
	</table>
	
	</div>
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
