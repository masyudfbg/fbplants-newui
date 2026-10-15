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
#detail-modal-overlay {
  display: none;
  position: fixed;
  inset: 0;
  background: rgba(0,0,0,0.55);
  z-index: 9999;
  align-items: center;
  justify-content: center;
}
#detail-modal-overlay.active {
  display: flex;
}
#detail-modal-box {
  background: #fff;
  border-radius: 8px;
  box-shadow: 0 8px 40px rgba(0,0,0,0.35);
  width: 92vw;
  max-width: 1300px;
  height: 85vh;
  display: flex;
  flex-direction: column;
  overflow: hidden;
  position: relative;
}
#detail-modal-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 10px 16px;
  background: #00A300;
  color: #fff;
  font-family: Arial, sans-serif;
  font-size: 14px;
  font-weight: bold;
  flex-shrink: 0;
}
#detail-modal-close {
  cursor: pointer;
  font-size: 20px;
  line-height: 1;
  background: none;
  border: none;
  color: #fff;
  padding: 0 4px;
}
#detail-modal-close:hover { opacity: 0.75; }
#detail-modal-iframe {
  flex: 1;
  width: 100%;
  border: none;
}
#detail-modal-loading {
  position: absolute;
  inset: 40px 0 0 0;
  display: flex;
  align-items: center;
  justify-content: center;
  font-family: Arial, sans-serif;
  color: #555;
  font-size: 14px;
  pointer-events: none;
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



	'if not rspt.eof then
	''	companyview = rspt("fccompanyname")
	''	While Not rspt.EOF
	''
	''		fcbaloop = rspt("fccode")
	''		'Response.Write fcbaloop 
	''		sqlexec = "BEGIN " & _ 
    ''         "opr_accountreport_bulanan_api.Proses('" & bulan & "','" & tahun & "','" & fcbaloop & "');" & _
	''		  "END; " 
    ''		connStr.Execute(sqlexec)

	''	rspt.MoveNext
	''	Wend
	''	rspt.close()
	''	rspt	=	null

	' END IF

	

	

	q = "Select region, fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & fcba & "' "
		' response.write q
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

	'rubah format tgl to jd dd/mm/yyyy
	bulansplitto  	= left(actto,2)
	tglsplitto    	= mid(actto,4,2)
	thnsplitto    	= right(actto,4)
	tglto     		= tglsplitto & "/" & bulansplitto & "/" & thnsplitto

	'tglawal tahun
	'tglawalsbi = "01" & "/" & "01" & "/" & thnsplitfrom


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
	 
	 	'filter by company
	if trim(company) = "all" then 
		filtercompany = " WHERE b.region = '" & region & "'"
	else
		if trim(fcba) = "-" then 
	    	filtercompany = " WHERE b.fccompanycode ='" & company & "' "
		else 
			filtercompany = " WHERE b.fccode ='" & fcba & "' "
		end if
	end if

	'bulan lalu
    'if trim(t1) = "Tgl" THEN
		filterprdlalu = " AND (o.transdate between to_date('" & prev_bulanfrom & "','mm/dd/yyyy') and to_date('" & prev_bulanto & "','mm/dd/yyyy')) " 			   
	'else
	'	filterprdlalu = " AND o.period_month='" & xbln2 & "' " & _ 
	'				" AND o.period_year ='" & xthn2 & "' " 		   
	'end if

    'if trim(t1) = "Tgl" THEN
		filterprd = " AND (o.transdate between to_date('" & actfrom & "','mm/dd/yyyy') and to_date('" & actto & "','mm/dd/yyyy')) " 			   
	'else
	'	filterprd = " AND o.period_month='" & bulan & "' " & _ 
	'				" AND o.period_year ='" & tahun & "' " 		   
	'end if

	'filter sbi
    'if trim(t1) = "Tgl" THEN
		filterprdsbi = " AND (o.transdate between to_date('" & tglawalsbi & "','mm/dd/yyyy') and to_date('" & actto & "','mm/dd/yyyy')) " 			   
	'else
	'	filterprdsbi = " AND o.period_month <= '" & bulan & "' AND o.period_year ='" & tahun & "' " 		   
	'end if

	'filter by cost elemen
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


    sqlview =  "select t.region,t.fccompanycode,t.fcba,t.fieldstatus,t.source,t.budgetcode,t.budgetname,t.jobgroupcode,t.jobgroupname,t.jobcode,t.jobname, " & _
	" t.sectioncode,t.fieldcode,NVL(t.fieldarea,0) fieldarea,NVL(t.fieldpokok,0) fieldpokok,jobmaster_api.get_uomunit(t.jobcode) uom," & _
	" NVL(sum(t.hkblnlalu),0)hkblnlalu,NVL(sum(t.fisikblnlalu),0)fisikblnlalu,NVL(sum(t.unitblnlalu),0)unitblnlalu,NVL(sum(t.amountblnlalu),0)amountblnlalu, " & _
	" sum(t.hkblnini)hkblnini,sum(t.fisikblnini)fisikblnini,sum(t.unitblnini)unitblnini,sum(t.amountblnini)amountblnini, " & _
	" sum(t.hksbi)hksbi,nvl(sum(t.fisiksbi),0)fisiksbi,sum(t.unitsbi)unitsbi,sum(t.amountsbi)amountsbi, " & _
	" NVL(SUM(t.rkb_hkblnlalu),0)rkb_hkblnlalu,NVL(SUM(t.rkb_outputblnlalu),0)rkb_outputblnlalu,NVL(SUM(t.rkb_amountblnlalu),0)rkb_amountblnlalu, " & _
	" NVL(SUM(t.rkb_hkblnini),0)rkb_hkblnini,NVL(SUM(t.rkb_outputblnini),0)rkb_outputblnini,NVL(SUM(t.rkb_amountblnini),0)rkb_amountblnini, " & _
	" NVL(SUM(t.rkb_hksbi),0)rkb_hksbi,NVL(SUM(t.rkb_outputsbi),0)rkb_outputsbi,NVL(SUM(t.rkb_amountsbi),0)rkb_amountsbi, " & _
	" COALESCE(NVL(SUM(t.amountblnlalu),0) / NULLIF(NVL(SUM(t.unitblnlalu),0),0),0) rphablnlalu, " & _
	" COALESCE(NVL(SUM(t.amountblnini),0) / NULLIF(NVL(SUM(t.unitblnini),0),0),0) rphablnini, " & _ 
	" COALESCE(NVL(SUM(t.amountsbi),0) / NULLIF(NVL(SUM(t.unitsbi),0),0),0) rphasbi," & _
	" COALESCE(NVL(SUM(t.rkb_amountblnlalu),0) / NULLIF(NVL(SUM(t.rkb_outputblnlalu),0),0),0) rkb_rphablnlalu," & _ 
	" COALESCE(NVL(SUM(t.rkb_amountblnini),0) / NULLIF(NVL(SUM(t.rkb_outputblnini),0),0),0) rkb_rphablnini," & _ 
	" COALESCE(NVL(SUM(t.rkb_amountsbi),0) / NULLIF(NVL(SUM(t.rkb_outputsbi),0),0),0) rkb_rphasbi," & _
	" DECODE(COALESCE(NVL(SUM(t.rkb_amountblnlalu),0) / NULLIF(NVL(SUM(t.rkb_outputblnlalu),0),0),0),0,1,COALESCE(NVL(SUM(t.amountblnlalu),0) / NULLIF(NVL(SUM(t.unitblnlalu),0),0),0) / COALESCE(NVL(SUM(t.rkb_amountblnlalu),0) / NULLIF(NVL(SUM(t.rkb_outputblnlalu),0),0),0))*100 var_blnlalu, " & _ 
	" DECODE(COALESCE(NVL(SUM(t.rkb_amountblnini),0) / NULLIF(NVL(SUM(t.rkb_outputblnini),0),0),0),0,1,COALESCE(NVL(SUM(t.amountblnini),0) / NULLIF(NVL(SUM(t.unitblnini),0),0),0) / COALESCE(NVL(SUM(t.rkb_amountblnini),0) / NULLIF(NVL(SUM(t.rkb_outputblnini),0),0),0))*100 var_blnini, " & _ 
	" DECODE(COALESCE(NVL(SUM(t.rkb_amountsbi),0) / NULLIF(NVL(SUM(t.rkb_outputsbi),0),0),0),0,1,COALESCE(NVL(SUM(t.amountsbi),0) / NULLIF(NVL(SUM(t.unitsbi),0),0),0) / COALESCE(NVL(SUM(t.rkb_amountsbi),0) / NULLIF(NVL(SUM(t.rkb_outputsbi),0),0),0))*100 var_sbi " & _ 
	" from( " & _
			" SELECT b.region,b.fccompanycode,b.fccode fcba,o.source,o.GROUPCODE  jobgroupcode, o.GROUPNAME  jobgroupname,o.fccode  jobcode, o.fcname  jobname, o.fccode_status fieldstatus, " & _
			" o.budgetcode, o.budgetname," & _
			" o.sectioncode,o.fieldcode,o.fieldarea,o.fieldpokok, " & _
			" decode("&bulan&",1,0,nvl(SUM(o.hk),0)) hkblnlalu, decode("&bulan&",1,0,nvl(SUM(o.OUTPUT),0)) fisikblnlalu, decode("&bulan&",1,0,nvl(SUM(o.unit),0)) unitblnlalu,decode("&bulan&",1,0,nvl(sum(o.amount),0)) amountblnlalu, " & _
			" 0 hkblnini,0 fisikblnini,0 unitblnini,0 amountblnini, " & _
			" 0 hksbi,0 fisiksbi,0 unitsbi,0 amountsbi, " & _
			" sum(o.rkb_hk) rkb_hkblnlalu,sum(o.rkb_output) rkb_outputblnlalu,sum(o.rkb_amount) rkb_amountblnlalu, " & _
			" 0 rkb_hkblnini,0 rkb_outputblnini,0 rkb_amountblnini, " & _
			" 0 rkb_hksbi,0 rkb_outputsbi,0 rkb_amountsbi " & _
			" FROM OPR_ACCOUNTREPORT_BULANAN  o " & _
			" LEFT JOIN BUSINESSUNIT b on o.FCBA=b.fccode " & _
			filtercompany & _
			filterprdlalu & _
			filtercostelemen & _
			filtertmtbm & _
			filterdiv & _
			filterCoa & _
			" group by  b.region,b.fccompanycode,b.fccode,o.source,o.GROUPCODE, o.GROUPNAME,o.fccode, o.fcname,o.fccode_status, o.sectioncode,o.fieldcode,o.fieldarea,o.fieldpokok,o.budgetcode,o.budgetname " & _
			"  union all " & _
			" SELECT b.region,b.fccompanycode,b.fccode fcba,o.source,o.GROUPCODE  jobgroupcode, o.GROUPNAME  jobgroupname,o.fccode  jobcode, o.fcname  jobname, o.fccode_status fieldstatus, " & _
			" o.budgetcode, o.budgetname," & _
			" o.sectioncode,o.fieldcode,o.fieldarea,o.fieldpokok, " & _
			" 0 hkblnlalu,0 fisikblnlalu,0 unitblnlalu,0 amountblnlalu, " & _
			" SUM(o.hk) hkblnini,SUM(o.OUTPUT) fisikblnini,SUM(o.unit) unitblnini,sum(o.amount)amountblnini, " & _
			" 0 hksbi,0 fisiksbi,0 unitsbi,0 amountsbi, " & _
			" 0 rkb_hkblnlalu,0 rkb_outputblnlalu,0 rkb_amountblnlalu, " & _
			" SUM(o.rkb_hk)rkb_hkblnini,SUM(o.rkb_output)rkb_outputblnini,SUM(o.rkb_amount)rkb_amountblnini, " & _
			" 0 rkb_hksbi,0 rkb_outputsbi,0 rkb_amountsbi " & _
			" FROM OPR_ACCOUNTREPORT_BULANAN  o " & _
			" LEFT JOIN BUSINESSUNIT b on o.FCBA=b.fccode " & _
			filtercompany & _
			filterprd & _
			filtercostelemen & _
			filtertmtbm & _
			filterdiv & _
			filterCoa & _
			" group by  b.region,b.fccompanycode,b.fccode,o.source,o.GROUPCODE, o.GROUPNAME,o.fccode, o.fcname,o.fccode_status, o.sectioncode,o.fieldcode,o.fieldarea,o.fieldpokok,o.budgetcode,o.budgetname " & _
			"  union all " & _
			" SELECT b.region,b.fccompanycode,b.fccode fcba,o.source,o.GROUPCODE, o.GROUPNAME,o.fccode  jobcode, o.fcname  jobname, o.fccode_status fieldstatus, " & _
			" o.budgetcode, o.budgetname," & _
			" o.sectioncode,o.fieldcode,o.fieldarea,o.fieldpokok, " & _
			" 0 hkblnlalu,0 fisikblnlalu,0 unitblnlalu,0 amountblnlalu, " & _
			" 0 hkblnini,0 fisikblnini,0 unitblnini,0 amountblnini, " & _
			" SUM(o.hk) hksbi,SUM(o.OUTPUT) fisiksbi,SUM(o.unit) unitsbi,sum(o.amount)amountsbi, " & _
			" 0 rkb_hkblnlalu,0 rkb_outputblnlalu,0 rkb_amountblnlalu, " & _
			" 0 rkb_hkblnini,0 rkb_outputblnini,0 rkb_amountblnini, " & _
			" sum(o.rkb_hk) rkb_hksbi,sum(o.rkb_output) rkb_outputsbi,sum(o.rkb_amount) rkb_amountsbi " & _
			" FROM OPR_ACCOUNTREPORT_BULANAN  o " & _
			" LEFT JOIN BUSINESSUNIT b on o.FCBA=b.fccode " & _
			filtercompany & _
			filterprdsbi & _
			filtercostelemen & _
			filtertmtbm & _
			filterdiv & _
			filterCoa & _
			" group by  b.region,b.fccompanycode,b.fccode,o.source,o.GROUPCODE, o.GROUPNAME,o.fccode, o.fcname,o.fccode_status, o.sectioncode,o.fieldcode,o.fieldarea,o.fieldpokok,o.budgetcode,o.budgetname " & _
	" )t  " & _
	" group by t.region,t.fccompanycode,t.fcba,t.source,t.budgetcode,t.budgetname,t.jobgroupcode,t.jobgroupname,t.jobcode,t.jobname,t.fieldstatus, t.sectioncode,t.fieldcode,t.fieldarea,t.fieldpokok " & _
	" order by t.region,t.fccompanycode,t.fcba,t.budgetcode,t.budgetname,t.jobgroupcode,t.jobgroupname,t.jobcode,t.jobname,t.sectioncode,t.fieldcode,t.fieldstatus,t.source"

	'Response.Write sqlview

	Set rsdata = Server.CreateObject("ADODB.Recordset")
	rsdata.CursorType = 3
	rsdata.CursorLocation = 3
	rsdata.Open sqlview, connstr
	
	%>
	<body onload="">
		<div id="detail-modal-overlay">
			<div id="detail-modal-box">
				<div id="detail-modal-header">
				<span id="detail-modal-title">Detail Biaya</span>
				<button id="detail-modal-close" title="Tutup">&times;</button>
				</div>
				<div id="detail-modal-loading">Memuat data...</div>
				<iframe id="detail-modal-iframe" src="about:blank"></iframe>
			</div>
		</div>
	<div id="wrapper">
	<!--#include file="../print_excel.asp"-->
	 
	<table width="100%" id="product-table" class="xthf-site" id="">
		 <thead>
        	 <tr>
				<th colspan="46" style=" background-color:#FFFFFF; color:#000000; border:none">
					<div id="header-wrapper">
					<div id="head_judul" style="float:left; text-align:left;">
							<h4><%=region%></h4>
							<h4>PT. <%=companyview%></h4>
							<%
								IF fcbaname <> "-" then 
									fcbaname = fcba
								ELSE 
									fcbaname = "ALL"
								END IF
							%>
							<h5><%=fcbaname%></h5>
						</div>
						<div id="subheader" class="centered" style="width:50%; margin:auto;">
							<h1 class="judul">Laporan Kebun - Detail</h1>
							<h3>Periode: <%=periode%></h5>
						</div>
						<div style="clear:both;"></div>
					</div>
				</th>
			</tr>
				<!--<div class="printhide">-->
			
		<tr style="background-color:#00A300;color:#fff;">
			<th colspan="13" rowspan=3>Lokasi</th>
		</tr>
		<tr style="background-color:#00A300;color:#fff;">
			<th colspan="11">Bulan Lalu</th>
			<th colspan="11">Bulan Ini</th>
			<th colspan="11">s/d Bulan Ini</th>
		</tr>
		<tr style="background-color:#00A300;color:#fff;">
			<th colspan="5">Realisasi</th>
			<th colspan="5">RKB</th>
			<th colspan="">Var</th>
			<th colspan="5">Realisasi</th>
			<th colspan="5">RKB</th>
			<th colspan="">Var</th>
			<th colspan="5">Realisasi</th>
			<th colspan="5">RKB</th>
			<th colspan="">Var</th>
		</tr>
		<tr style="background-color:#00A300;color:#fff;">
		    
		    <th>Region</th>
			<th>Company</th>
			<th>FCBA</th>
			<th>Status</th>
			<th>Budget</th>
			<th>Group</th>
			<th>JobCode</th>
			<th>UOM</th>
			<th>CostElement</th>
			<th>Afdeling</th>
			<th>Blok</th>
			<th>Luas</th>
			<th>Pkk</th>
			
			<th>Hk</th>
			<th>Unit</th>
			<th>Output</th>
			<th>Rupiah</th>
			<th>Rp/Unit</th>
			
			<th>Hk</th>
			<th>Unit</th>
			<th>Output</th>
			<th>Rupiah</th>
			<th>Rp/Unit</th>
			<th>%</th>
			
			<th>Hk</th>
			<th>Unit</th>
			<th>Output</th>
			<th>Rupiah</th>
			<th>Rp/Unit</th>
			
			<th>Hk</th>
			<th>Unit</th>
			<th>Output</th>
			<th>Rupiah</th>
			<th>Rp/Unit</th>
			<th>%</th>
			
			<th>Hk</th>
			<th>Unit</th>
			<th>Output</th>
			<th>Rupiah</th>
			<th>Rp/Unit</th>
			
			<th>Hk</th>
			<th>Unit</th>
			<th>Output</th>
			<th>Rupiah</th>
			<th>Rp/Unit</th>
			<th>%</th>
		</tr>
		</thead>
		<%
				if Not rsdata.eof then
					nCountofField = rsdata.fields.count-1
				end if
				
				hattl = 0
				pkkttl = 0
				
				hkblnlaluttl=0
				unitblnlaluttl=0
				fisikblnlaluttl=0
				amountblnlaluttl=0
				rphablnlaluttl=0
				
				rkb_hkblnlaluttl=0
				rkb_unitblnlaluttl=0
				rkb_outputblnlaluttl=0
				rkb_amountblnlaluttl=0
				rkbhablnlaluttl=0
				
				hkblninittl=0
				unitblninittl=0
				fisikblninittl=0
				amountblninittl=0
				rphablninittl=0
				
				rkb_hkblninittl=0
				rkb_unitblninittl=0
				rkb_outputblninittl=0
				rkb_amountblninittl=0
				rkbhablninittl=0
				
				hksbittl=0
				unitsbittl=0
				fisiksbittl=0
				amountsbittl=0
				rphasbittl=0
				
				rkb_hksbittl=0
				rkb_unitsbittl=0
				rkb_outputsbittl=0
				rkb_amountsbittl=0
				rkbhasbittl=0
				
				While Not rsdata.EOF
					

		%>
		<tr class=''>
			<td ><%=rsdata("region")%></td>
			<td><%=rsdata("fccompanycode")%></td>
			<td><%=rsdata("fcba")%></td>
			<td><%=rsdata("fieldstatus")%></td>
			<td><%=rsdata("budgetcode") & " - " & rsdata("budgetname")%></td>
			<td><%=rsdata("jobgroupcode") & " - " & rsdata("jobgroupname")%></td>
			<td><%=rsdata("jobcode") & " - " & rsdata("jobname")%></td>
			<td><%=rsdata("uom") %></td>
			<td><%=rsdata("source")%></td>
			<td><%=rsdata("sectioncode")%></td>
			<td><%=rsdata("fieldcode")%></td>
			<td  align="right"><%=FbpNum2(rsdata("fieldarea"))%></td>
			<td align="right"><%=FbpNum2(rsdata("fieldpokok"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("hkblnlalu"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("unitblnlalu"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("fisikblnlalu"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("amountblnlalu"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rphablnlalu"),2)%></td>
			
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_hkblnlalu"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_outputblnlalu"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_outputblnlalu"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_amountblnlalu"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_rphablnlalu"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("var_blnlalu")) %></td>
			
			<td class="ratakanan"><%=FormatNumber(rsdata("hkblnini"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("unitblnini"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("fisikblnini"))%></td>
			<td class="ratakanan">
				<%=FormatNumber(rsdata("amountblnini"))%><br>
				<%
					bulansplitfrom1  = left(actfrom,2)
					tglsplitfrom1    = mid(actfrom,4,2)
					thnsplitfrom1    = right(actfrom,4)
					tglfrom1     	= bulansplitfrom & "/" & tglsplitfrom1 & "/" & thnsplitfrom1

					'rubah format tgl to jd dd/mm/yyyy
					bulansplitto1  	= left(actto,2)
					tglsplitto1    	= mid(actto,4,2)
					thnsplitto1    	= right(actto,4)
					tglto1     		= bulansplitto1 & "/" & tglsplitto1 & "/" & thnsplitto1
				%>
				<%if FormatNumber(rsdata("amountblnini")) > 0 then%>
					<a target="_blank" href="http://www.firstborneo.com/p_listing_lhd.asp?pcompany=<%=rsdata("fccompanycode")%>&bkmfrom=<%=tglfrom1%>&bkmto=<%=tglto1%>&pfcba=<%=rsdata("fcba")%>&idview=26&divisi=all&jobcode=<%=rsdata("jobcode")%>&blok=<%=rsdata("fieldcode")%>"
					style="color: blue">
						[Detail]
					</a>
				<%end if%>
			</td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rphablnini"))%></td>
			
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_hkblnini"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_outputblnini"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_outputblnini"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_amountblnini"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_rphablnini"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("var_blnini")) %></td>
			
			<td class="ratakanan"><%=FormatNumber(rsdata("hksbi"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("unitsbi"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("fisiksbi"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("amountsbi"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rphasbi"))%></td>
			
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_hksbi"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_outputsbi"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_outputsbi"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_amountsbi"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("rkb_rphasbi"))%></td>
			<td class="ratakanan"><%=FormatNumber(rsdata("var_sbi")) %></td>
		</tr>
		<%
		    hattl               = cdbl(hattl) + cdbl(rsdata("fieldarea"))
			pkkttl              = cdbl(pkkttl) + cdbl(rsdata("fieldpokok"))
			hkblnlaluttl		= cdbl(hkblnlaluttl) + cdbl(rsdata("hkblnlalu"))
			unitblnlaluttl		= cdbl(unitblnlaluttl) + cdbl(rsdata("unitblnlalu"))
			fisikblnlaluttl		= cdbl(fisikblnlaluttl) + cdbl(rsdata("fisikblnlalu"))
			amountblnlaluttl	= cdbl(amountblnlaluttl) + cdbl(rsdata("amountblnlalu"))
			rphablnlaluttl		= cdbl(rphablnlaluttl) + cdbl(rsdata("rphablnlalu"))

			rkb_hkblnlaluttl    = cdbl(rkb_hkblnlaluttl) + cdbl(rsdata("rkb_hkblnlalu"))
			rkb_unitblnlaluttl  = cdbl(rkb_unitblnlaluttl) + cdbl(rsdata("rkb_outputblnlalu"))
			rkb_outputblnlaluttl= cdbl(rkb_outputblnlaluttl) + cdbl(rsdata("rkb_outputblnlalu"))
			rkb_amountblnlaluttl= cdbl(rkb_amountblnlaluttl) + cdbl(rsdata("rkb_amountblnlalu"))
			rkbhablnlaluttl		= cdbl(rkbhablnlaluttl)	+ cdbl(rsdata("rkb_rphablnlalu"))

			hkblninittl			= cdbl(hkblninittl) + cdbl(rsdata("hkblnini"))
			unitblninittl		= cdbl(unitblninittl) + cdbl(rsdata("unitblnini"))
			fisikblninittl		= cdbl(fisikblninittl) + cdbl(rsdata("fisikblnini"))
			amountblninittl		= cdbl(amountblninittl) + cdbl(rsdata("amountblnini"))
			rphablninittl		= cdbl(rphablninittl) + cdbl(rsdata("rphablnini"))

			rkb_hkblninittl 	= cdbl(rkb_hkblninittl) + cdbl(rsdata("rkb_hkblnini"))
			rkb_unitblninittl 	= cdbl(rkb_unitblninittl) + cdbl(rsdata("rkb_outputblnini"))
			rkb_outputblninittl = cdbl(rkb_outputblninittl) + cdbl(rsdata("rkb_outputblnini"))
			rkb_amountblninittl = cdbl(rkb_amountblninittl) + cdbl(rsdata("rkb_amountblnini"))
			rkbhablninittl		= cdbl(rkbhablninittl) + cdbl(rsdata("rkb_rphablnini"))

			hksbittl			= cdbl(hksbittl) + cdbl(rsdata("hksbi"))
			unitsbittl			= cdbl(unitsbittl) + cdbl(rsdata("unitsbi"))
			fisiksbittl			= cdbl(fisiksbittl) + cdbl(rsdata("fisiksbi"))
			amountsbittl		= cdbl(amountsbittl) + cdbl(rsdata("amountsbi"))
			rphasbittl			= cdbl(rphasbittl) + cdbl(rsdata("rphasbi"))

			rkb_hksbittl		= cdbl(rkb_hksbittl) + cdbl(rsdata("rkb_hksbi"))
			rkb_unitsbittl		= cdbl(rkb_unitsbittl) + cdbl(rsdata("rkb_outputsbi"))
			rkb_outputsbittl	= cdbl(rkb_outputsbittl) + cdbl(rsdata("rkb_outputsbi"))
			rkb_amountsbittl	= cdbl(rkb_amountsbittl) + cdbl(rsdata("rkb_amountsbi"))
			rkbhasbittl			= cdbl(rkbhasbittl) + cdbl(rsdata("rkb_rphasbi"))

			rsdata.MoveNext
			Wend
			rsdata.close()
			rsdata	=	null
		%>
		<tr class="huruftebal ">
			<td colspan='11'>Total</td>
			<td class="ratakanan"><span id='ttlrkpfisikblnlalu'><%=FormatNumber(hattl)%></span></td>
			<td class="ratakanan"><span id='ttlrkpfisikblnlalu'><%=FormatNumber(pkkttl)%></span></td>
			
			<td class="ratakanan"><span id='ttlrkpfisikblnlalu'><%=FormatNumber(hkblnlaluttl)%></span></td>
			<td class="ratakanan"><span id='ttlrkpfisikblnlalu'><%=FormatNumber(unitblnlaluttl)%></span></td>
			<td class="ratakanan"><span id='ttlrkpfisikblnlalu'><%=FormatNumber(fisikblnlaluttl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprpblnlalu'><%=FormatNumber(amountblnlaluttl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprphablnlalu'><%=FormatNumber(rphablnlaluttl)%></span></td>
			
			<td class="ratakanan"><span id='ttlrkprkbfisikblnlalu'><%=FormatNumber(rkb_hkblnlaluttl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprkbfisikblnlalu'><%=FormatNumber(rkb_unitblnlaluttl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprkbfisikblnlalu'><%=FormatNumber(rkb_outputblnlaluttl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprkbrpblnlalu'><%=FormatNumber(rkb_amountblnlaluttl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprkbrphablnlalu'><%=FormatNumber(rkbhablnlaluttl)%></span></td>
			<td class="ratakanan"><span id='ttlrkpvarianblnlalu'>0</span></td>

			<td class="ratakanan"><span id='ttlrkpfisikblnini'><%=FormatNumber(hkblninittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkpfisikblnini'><%=FormatNumber(unitblninittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkpfisikblnini'><%=FormatNumber(fisikblninittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprpblnini'><%=FormatNumber(amountblninittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprphablnini'><%=FormatNumber(rphablninittl)%></span></td>
			
			<td class="ratakanan"><span id='ttlrkprkbfisikblnini'><%=FormatNumber(rkb_hkblninittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprkbfisikblnini'><%=FormatNumber(rkb_unitblninittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprkbfisikblnini'><%=FormatNumber(rkb_outputblninittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprkbrpblnini'><%=FormatNumber(rkb_amountblninittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprkbrphablnini'><%=FormatNumber(rkbhablninittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkpvarianblnini'>0</span></td>

			<td class="ratakanan"><span id='ttlrkpfisikblnsbi'><%=FormatNumber(hksbittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkpfisikblnsbi'><%=FormatNumber(unitsbittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkpfisikblnsbi'><%=FormatNumber(fisiksbittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprpblnsbi'><%=FormatNumber(amountsbittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprphablnsbi'><%=FormatNumber(rphasbittl)%></span></td>
			
			<td class="ratakanan"><span id='ttlrkprkbfisikblnsbi'><%=FormatNumber(rkb_hksbittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprkbfisikblnsbi'><%=FormatNumber(rkb_unitsbittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprkbfisikblnsbi'><%=FormatNumber(rkb_outputsbittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprkbrpblnsbi'><%=FormatNumber(rkb_amountsbittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkprkbrphablnsbi'><%=FormatNumber(rkbhasbittl)%></span></td>
			<td class="ratakanan"><span id='ttlrkpvarianblnsbi'>0</span></td>
		</tr>
	</table>
	</div>
 </tbody>	

<!--#include file="bottom.Asp"-->
