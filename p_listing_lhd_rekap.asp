<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting -LHD Listing</title>
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

Dim tbl,doc,filterdiv,filterlhd,pcompany,fcba,bkmfrom,bkmto,divisi,bulansplitfrom,tglsplitfrom,thnsplitfrom,tglfrom,tglTo,blnthn,q
dim sqldate,rsdate,startdate,sqlsbi,rssbi,startdatesbi
cmdstyle="border-right: 1 solid #000000; border-bottom: 1 solid #000000"
cmdstyleTop="border-right: 1 solid #000000; border-bottom: 1 solid #000000; border-top: 1 solid #000000"
cmdstyleLeft="border-right: 1 solid #000000; border-bottom: 1 solid #000000; border-Left: 1 solid #000000"




    pcompany 	= Request("pcompany")
    fcba 		= Request("pfcba")
    bkmfrom 	= Request("bkmfrom")
    bkmto   	= Request("bkmto")
    divisi   	= Request("divisi")
	jobcode		= Request("jobcode")
	blok		= Request("blok")


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

	bulansplitfrom  = left(bkmfrom,2)
	tglsplitfrom    = mid(bkmfrom,4,2)
	thnsplitfrom    = right(bkmfrom,4)

	bulansplitTo  = left(bkmto,2)
	tglsplitTo    = mid(bkmto,4,2)
	thnsplitTo    = right(bkmto,4)


	tglfrom     	= bulansplitfrom & "/" & tglsplitfrom & "/" & thnsplitfrom
	tglTo			= bulansplitTo & "/" & tglsplitTo & "/" & thnsplitTo
	blnthn 			= bulansplitfrom&thnsplitfrom



	sqldate		= " select start_date from empalldedstatus where fcba='" & fcba & "' and to_date('" & bkmfrom & "','mm/dd/yyyy') between  start_date and end_date"
	set rsdate 	= connStr.Execute(sqldate)
	if not rsdate.eof then
		startdate = rsdate("start_date")
	end if

	'startdate sbi
	sqlsbi		= " select start_date from empalldedstatus where fcba='" & fcba & "' and year='" & thnsplitfrom & "' and month='1'"
	set rssbi 	= connStr.Execute(sqlsbi)
	if not rssbi.eof then
		startdatesbi = rssbi("start_date")
	end if



	sqlexec = "BEGIN " & _ 
              "opr_accountreport_bulanan_api.Proses('" & bulansplitfrom & "','" & thnsplitfrom & "','" & fcba & "');" & _
			  "END; " 
    connStr.Execute(sqlexec)


	if trim(divisi) <> "all" then
		filterdiv = " AND t.sectioncode='" & divisi & "'"
	end if

	if trim(jobcode) <> "" then
		filterfccode = " AND t.fccode='" & jobcode & "'"
	end if

	if trim(blok) = "null" then
		filterBlok = ""
	else
		filterBlok = " AND t.fieldcode ='" & blok & "'"
	end if


	if trim(fcba) = "-" then 
		filterlhd = " WHERE t.FCBA in (SELECT fccode from businessunit where fccompanycode='" & pcompany & "') "
	else
	    filterlhd = " WHERE t.FCBA ='" & fcba & "' "
	end if


sqlview1 = "SELECT t.fcba,t.fccode,t.fcname,t.sectioncode,t.fieldcode,t.fieldarea,t.fieldstatus,t.uom_unit, " & _
" nvl(sum(t.fisik),0)fisik,nvl(sum(t.fisikshi),0)fisikshi,nvl(sum(t.fisiksbi),0)fisiksbi," & _
" nvl(sum(t.unit),0)unit,nvl(sum(t.unitshi),0)unitshi,nvl(sum(t.unitsbi),0)unitsbi, " & _
" nvl(sum(t.hk),0) hk,nvl(sum(t.hkshi),0) hkshi,nvl(sum(t.hksbi),0) hksbi, " & _
" nvl(sum(t.rphkhi),0)rphkhi,nvl(sum(t.rphkshi),0)rphkshi,nvl(sum(t.rphksbi),0)rphksbi, " & _
" nvl(sum(t.fisikbrhi),0)fisikbrhi,nvl(sum(t.fisikbrshi),0)fisikbrshi,nvl(sum(t.fisikbrsbi),0)fisikbrsbi, " & _
" nvl(sum(t.unitbrhi),0)unitbrhi,nvl(sum(t.unitbrshi),0)unitbrshi,nvl(sum(t.unitbrsbi),0)unitbrsbi, " & _
" nvl(sum(t.rpbrhi),0)rpbrhi,nvl(sum(t.rpbrshi),0)rpbrshi,nvl(sum(t.rpbrsbi),0)rpbrsbi, " & _
" t.itembahan,nvl(sum(t.qtybahanhi),0)qtybahanhi,nvl(sum(t.qtybahanshi),0)qtybahanshi,nvl(sum(t.qtybahansbi),0)qtybahansbi, " & _
" nvl(sum(t.biayabahanhi),0)biayabahanhi,nvl(sum(t.biayabahanshi),0)biayabahanshi,nvl(sum(t.biayabahansbi),0)biayabahansbi, " & _
" t.itemkendaraan,nvl(sum(t.qtyvrahi),0)qtyvrahi,nvl(sum(t.qtyvrashi),0)qtyvrashi,nvl(sum(t.qtyvrasbi),0)qtyvrasbi, " & _
" nvl(sum(t.biayavrahi),0)biayavrahi,nvl(sum(t.biayavrashi),0)biayavrashi,nvl(sum(t.biayavrasbi),0)biayavrasbi" & _
" FROM ( " & _
"        SELECT o.fcba,o.fccode,o.fcname,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus,o.uom_unit, " & _
"        SUM(o.output) fisik,0 fisikshi,0 fisiksbi,SUM(o.unit) unit,0 unitshi,0 unitsbi, " & _
"        SUM(o.hk) hk,0 hkshi,0 hksbi, round(SUM(o.amount),0) rphkhi, 0 rphkshi,0 rphksbi, " & _
"        0 fisikbrhi,0 fisikbrshi,0 fisikbrsbi, " & _
"        0 unitbrhi,0 unitbrshi,0 unitbrsbi, " & _
"        0 rpbrhi,0 rpbrshi,0 rpbrsbi, " & _
"		 o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom itembahan,0 qtybahanhi,0 qtybahanshi,0 qtybahansbi,0 biayabahanhi,0 biayabahanshi,0 biayabahansbi, " & _
"		 o.vracode || ' - ' || o.vraname itemkendaraan,0 qtyvrahi,0 qtyvrashi,0 qtyvrasbi,0 biayavrahi,0 biayavrashi,0 biayavrasbi  " & _
"        FROM OPR_ACCOUNTREPORT_BULANAN o " & _
"        where o.transdate=to_date('" & bkmfrom & "','mm/dd/yyyy') AND o.source='1.UPAH' " & _ 
"        GROUP BY o.fcba,o.fccode,o.fcname,o.uom_unit,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus,o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom,o.vracode || ' - ' || o.vraname " & _

"        union all" & _
"        SELECT o.fcba,o.fccode,o.fcname,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus,o.uom_unit, " & _
"        0 fisik,SUM(o.output) fisikshi,0 fisiksbi,0 unit,SUM(o.unit) unitshi,0 unitsbi, " & _
"        0 hk,SUM(o.hk) hkshi,0 hksbi, 0 rphkhi,round(SUM(o.amount),0) rphkshi,0 rphksbi, " & _
"        0 fisikbrhi,0 fisikbrshi,0 fisikbrsbi, " & _
"        0 unitbrhi,0 unitbrshi,0 unitbrsbi, " & _
"        0 rpbrhi,0 rpbrshi,0 rpbrsbi,  " & _
"		 o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom itembahan,0 qtybahanhi,0 qtybahanshi,0 qtybahansbi,0 biayabahanhi,0 biayabahanshi,0 biayabahansbi, " & _
"		 o.vracode || ' - ' || o.vraname itemkendaraan,0 qtyvrahi,0 qtyvrashi,0 qtyvrasbi,0 biayavrahi,0 biayavrashi,0 biayavrasbi " & _
"        FROM OPR_ACCOUNTREPORT_BULANAN o " & _
"        where o.transdate between to_date('" & bkmfrom & "','mm/dd/yyyy') AND to_date('" & bkmto & "','mm/dd/yyyy') AND o.source='1.UPAH' " & _ 
"        GROUP BY o.fcba,o.fccode,o.fcname,o.uom_unit,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus,o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom,o.vracode || ' - ' || o.vraname " & _

"        union all " & _
"        SELECT o.fcba,o.fccode,o.fcname,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus,o.uom_unit,  " & _
"        0 fisik,0 fisikshi,SUM(o.output) fisiksbi,0 unit,0 unitshi,SUM(o.unit) unitsbi, " & _
"        0 hk,0 hkshi,SUM(o.hk) hksbi, 0 rphkhi,0 rphkshi,round(SUM(o.amount),0) rphksbi, " & _
"        0 fisikbrhi,0 fisikbrshi,0 fisikbrsbi, " & _
"        0 unitbrhi,0 unitbrshi,0 unitbrsbi, " & _
"        0 rpbrhi,0 rpbrshi,0 rpbrsbi, " & _
"		 o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom itembahan,0 qtybahanhi,0 qtybahanshi,0 qtybahansbi,0 biayabahanhi,0 biayabahanshi,0 biayabahansbi, " & _
"		 o.vracode || ' - ' || o.vraname itemkendaraan,0 qtyvrahi,0 qtyvrashi,0 qtyvrasbi,0 biayavrahi,0 biayavrashi,0 biayavrasbi " & _
"        FROM OPR_ACCOUNTREPORT_BULANAN o " & _
"        where to_char(o.transdate,'mm')<='" & bulansplitfrom & "' and to_char(o.transdate,'yyyy')='" & thnsplitfrom & "' AND o.source='1.UPAH' " & _ 
"        GROUP BY o.fcba,o.fccode,o.fcname,o.uom_unit,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus,o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom,o.vracode || ' - ' || o.vraname " & _

"        union all " & _
"        SELECT o.fcba,o.fccode,o.fcname,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus,o.uom_unit, " & _
"        0 fisik,0 fisikshi,0 fisiksbi,0 unit,0 unitshi,0 unitsbi, " & _
"        0 hk,0 hkshi,0 hksbi, 0 rphkhi, 0 rphkshi,0 rphksbi, " & _
"        SUM(o.output) fisikbrhi,0 fisikbrshi,0 fisikbrsbi, " & _
"        SUM(o.unit) unitbrhi,0 unitbrshi,0 unitbrsbi, " & _
"        round(SUM(o.amount),0) rpbrhi,0 rpbrshi,0 rpbrsbi, " & _
"		 o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom itembahan,0 qtybahanhi,0 qtybahanshi,0 qtybahansbi,0 biayabahanhi,0 biayabahanshi,0 biayabahansbi, " & _
"		 o.vracode || ' - ' || o.vraname itemkendaraan,0 qtyvrahi,0 qtyvrashi,0 qtyvrasbi,0 biayavrahi,0 biayavrashi,0 biayavrasbi " & _
"        FROM OPR_ACCOUNTREPORT_BULANAN o " & _
"        where o.transdate=to_date('" & bkmfrom & "','mm/dd/yyyy') AND o.source='4.WIP' " & _ 
"        GROUP BY o.fcba,o.fccode,o.fcname,o.uom_unit,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus,o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom,o.vracode || ' - ' || o.vraname " & _

"        union all " & _
"        SELECT o.fcba,o.fccode,o.fcname,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus,o.uom_unit, " & _
"        0 fisik,0 fisikshi,0 fisiksbi,0 unit,0 unitshi,0 unitsbi, " & _
"        0 hk,0 hkshi,0 hksbi, 0 rphkhi,0 rphkshi,0 rphksbi, " & _
"        0 fisikbrhi,SUM(o.output) fisikbrshi,0 fisikbrsbi, " & _
"        0 unitbrhi,SUM(o.unit) unitbrshi,0 unitbrsbi, " & _
"        0 rpbrhi,round(SUM(o.amount),0) rpbrshi,0 rpbrsbi, " & _
"		 o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom itembahan,0 qtybahanhi,0 qtybahanshi,0 qtybahansbi,0 biayabahanhi,0 biayabahanshi,0 biayabahansbi, " & _
"		 o.vracode || ' - ' || o.vraname itemkendaraan,0 qtyvrahi,0 qtyvrashi,0 qtyvrasbi,0 biayavrahi,0 biayavrashi,0 biayavrasbi " & _
"        FROM OPR_ACCOUNTREPORT_BULANAN o " & _
"        where o.transdate between to_date('" & bkmfrom & "','mm/dd/yyyy') AND to_date('" & bkmto & "','mm/dd/yyyy') AND o.source='4.WIP' " & _ 
"        GROUP BY o.fcba,o.fccode,o.fcname,o.uom_unit,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus,o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom,o.vracode || ' - ' || o.vraname " & _

"        union all " & _
"        SELECT o.fcba,o.fccode,o.fcname,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus,o.uom_unit," & _
"        0 fisik,0 fisikshi,0 fisiksbi,0 unit,0 unitshi,0 unitsbi, " & _
"        0 hk,0 hkshi,0 hksbi, 0 rphkhi,0 rphkshi,0 rphksbi, " & _
"        0 fisikbrhi,0 fisikbrshi,SUM(o.output) fisikbrsbi, " & _
"        0 unitbrhi,0 unitbrshi,SUM(o.unit) unitbrsbi, " & _
"        0 rpbrhi,0 rpbrshi,round(SUM(o.amount),0) rpbrsbi, " & _
"		 o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom itembahan,0 qtybahanhi,0 qtybahanshi,0 qtybahansbi,0 biayabahanhi,0 biayabahanshi,0 biayabahansbi, " & _
"		 o.vracode || ' - ' || o.vraname itemkendaraan,0 qtyvrahi,0 qtyvrashi,0 qtyvrasbi,0 biayavrahi,0 biayavrashi,0 biayavrasbi " & _
"        FROM OPR_ACCOUNTREPORT_BULANAN o " & _
"        where to_char(o.transdate,'mm')<='" & bulansplitfrom & "' and to_char(o.transdate,'yyyy')='" & thnsplitfrom & "' AND o.source='4.WIP' " & _ 
"        GROUP BY o.fcba,o.fccode,o.fcname,o.uom_unit,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus,o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom,o.vracode || ' - ' || o.vraname " & _

"	    union all " & _ 
"		SELECT o.fcba,o.fccode,o.fcname,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus,o.uom_unit, " & _ 
"		0 fisik,0 fisikshi,0 fisiksbi,0 unit,0 unitshi,0 unitsbi, " & _ 
"		0 hk,0 hkshi,0 hksbi, 0 rphkhi, 0 rphkshi,0 rphksbi, 0 fisikbrhi,0 fisikbrshi, " & _ 
"		0 fisikbrsbi, 0 unitbrhi,0 unitbrshi,0 unitbrsbi, 0 rpbrhi,0 rpbrshi,0 rpbrsbi, " & _ 
"		o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom itembahan,sum(o.unit) qtybahanhi,0 qtybahanshi,0 qtybahansbi, round(sum(o.amount),0) biayabahanhi,0 biayabahanshi,0 biayabahansbi,   " & _
"		 o.vracode || ' - ' || o.vraname itemkendaraan,0 qtyvrahi,0 qtyvrashi,0 qtyvrasbi,0 biayavrahi,0 biayavrashi,0 biayavrasbi " & _ 
"		FROM OPR_ACCOUNTREPORT_BULANAN o " & _ 
"		where o.transdate=to_date('" & bkmfrom & "','mm/dd/yyyy') AND o.source='2.BAHAN' " & _ 
"		GROUP BY o.fcba,o.fccode,o.fcname,o.uom_unit,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus, o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom,o.vracode || ' - ' || o.vraname " & _ 

"	    union all  " & _ 
"		SELECT o.fcba,o.fccode,o.fcname,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus,o.uom_unit, " & _ 
"		0 fisik,0 fisikshi,0 fisiksbi,0 unit,0 unitshi,0 unitsbi, " & _ 
"		0 hk,0 hkshi,0 hksbi, 0 rphkhi, 0 rphkshi,0 rphksbi, 0 fisikbrhi,0 fisikbrshi, " & _ 
"		0 fisikbrsbi, 0 unitbrhi,0 unitbrshi,0 unitbrsbi, 0 rpbrhi,0 rpbrshi,0 rpbrsbi, " & _ 
"		o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom itembahan,0 qtybahanhi,sum(o.unit) qtybahanshi,0 qtybahansbi,0 biayabahanhi,round(sum(o.amount),0) biayabahanshi,0 biayabahansbi,   " & _ 
"		 o.vracode || ' - ' || o.vraname itemkendaraan,0 qtyvrahi,0 qtyvrashi,0 qtyvrasbi,0 biayavrahi,0 biayavrashi,0 biayavrasbi " & _
"		FROM OPR_ACCOUNTREPORT_BULANAN o " & _ 
"		where o.transdate between to_date('" & bkmfrom & "','mm/dd/yyyy') AND to_date('" & bkmto & "','mm/dd/yyyy') AND o.source='2.BAHAN' " & _ 
"		GROUP BY o.fcba,o.fccode,o.fcname,o.uom_unit,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus, o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom,o.vracode || ' - ' || o.vraname " & _ 

"	    union all  " & _ 
"		SELECT o.fcba,o.fccode,o.fcname,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus,o.uom_unit, " & _ 
"		0 fisik,0 fisikshi,0 fisiksbi,0 unit,0 unitshi,0 unitsbi, " & _ 
"		0 hk,0 hkshi,0 hksbi, 0 rphkhi, 0 rphkshi,0 rphksbi, 0 fisikbrhi,0 fisikbrshi, " & _ 
"		0 fisikbrsbi, 0 unitbrhi,0 unitbrshi,0 unitbrsbi, 0 rpbrhi,0 rpbrshi,0 rpbrsbi, " & _ 
"		o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom itembahan,0 qtybahanhi,0 qtybahanshi,sum(o.unit) qtybahansbi,0 biayabahanhi,0 biayabahanshi,round(sum(o.amount),0) biayabahansbi,   " & _ 
"		o.vracode || ' - ' || o.vraname itemkendaraan,0 qtyvrahi,0 qtyvrashi,0 qtyvrasbi,0 biayavrahi,0 biayavrashi,0 biayavrasbi " & _
"		FROM OPR_ACCOUNTREPORT_BULANAN o " & _ 
"		where to_char(o.transdate,'mm')<='" & bulansplitfrom & "' and to_char(o.transdate,'yyyy')='" & thnsplitfrom & "' AND o.source = '2.BAHAN' " & _ 
"		GROUP BY o.fcba,o.fccode,o.fcname,o.uom_unit,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus, o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom,o.vracode || ' - ' || o.vraname " & _ 

"		union all  " & _ 
"		SELECT o.fcba,o.fccode,o.fcname,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus,o.uom_unit, " & _ 
"		0 fisik,0 fisikshi,0 fisiksbi,0 unit,0 unitshi,0 unitsbi, " & _ 
"		0 hk,0 hkshi,0 hksbi, 0 rphkhi, 0 rphkshi,0 rphksbi, 0 fisikbrhi,0 fisikbrshi, " & _ 
"		0 fisikbrsbi, 0 unitbrhi,0 unitbrshi,0 unitbrsbi, 0 rpbrhi,0 rpbrshi,0 rpbrsbi, " & _ 
"		o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom itembahan,0 qtybahanhi,0 qtybahanshi,0 qtybahansbi,0 biayabahanhi,0 biayabahanshi,0 biayabahansbi, " & _ 
"		o.vracode || ' - ' || o.vraname itemkendaraan,SUM(o.unit) qtyvrahi,0 qtyvrashi,0 qtyvrasbi,round(sum(o.amount),0) biayavrahi,0 biayavrashi,0 biayavrasbi   " & _ 
"		FROM OPR_ACCOUNTREPORT_BULANAN o " & _ 
"		where o.transdate=to_date('" & bkmfrom & "','mm/dd/yyyy') AND o.source='3.VRA' " & _ 
"		GROUP BY o.fcba,o.fccode,o.fcname,o.uom_unit,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus, o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom,o.vracode || ' - ' || o.vraname " & _

"	    union all  " & _ 
"		SELECT o.fcba,o.fccode,o.fcname,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus,o.uom_unit, " & _ 
"		0 fisik,0 fisikshi,0 fisiksbi,0 unit,0 unitshi,0 unitsbi, " & _ 
"		0 hk,0 hkshi,0 hksbi, 0 rphkhi, 0 rphkshi,0 rphksbi, 0 fisikbrhi,0 fisikbrshi, " & _ 
"		0 fisikbrsbi, 0 unitbrhi,0 unitbrshi,0 unitbrsbi, 0 rpbrhi,0 rpbrshi,0 rpbrsbi, " & _ 
"		o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom itembahan,0 qtybahanhi,0 qtybahanshi,0 qtybahansbi,0 biayabahanhi,0 biayabahanshi,0 biayabahansbi, " & _ 
"		o.vracode || ' - ' || o.vraname itemkendaraan,0 qtyvrahi,SUM(o.unit) qtyvrashi,0 qtyvrasbi,0 biayavrahi,round(sum(o.amount),0) biayavrashi,0 biayavrasbi   " & _ 
"		FROM OPR_ACCOUNTREPORT_BULANAN o " & _ 
"		where o.transdate between to_date('" & bkmfrom & "','mm/dd/yyyy') AND to_date('" & bkmto & "','mm/dd/yyyy') AND o.source='3.VRA' " & _ 
"		GROUP BY o.fcba,o.fccode,o.fcname,o.uom_unit,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus, o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom,o.vracode || ' - ' || o.vraname " & _

"	    union all  " & _ 
"		SELECT o.fcba,o.fccode,o.fcname,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus,o.uom_unit, " & _ 
"		0 fisik,0 fisikshi,0 fisiksbi,0 unit,0 unitshi,0 unitsbi, " & _ 
"		0 hk,0 hkshi,0 hksbi, 0 rphkhi, 0 rphkshi,0 rphksbi, 0 fisikbrhi,0 fisikbrshi, " & _ 
"		0 fisikbrsbi, 0 unitbrhi,0 unitbrshi,0 unitbrsbi, 0 rpbrhi,0 rpbrshi,0 rpbrsbi, " & _ 
"		o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom itembahan,0 qtybahanhi,0 qtybahanshi,0 qtybahansbi,0 biayabahanhi,0 biayabahanshi,0 biayabahansbi, " & _ 
"		o.vracode || ' - ' || o.vraname itemkendaraan,0 qtyvrahi,0 qtyvrashi,SUM(o.unit) qtyvrasbi,0 biayavrahi,0 biayavrashi,round(sum(o.amount),0) biayavrasbi   " & _ 
"		FROM OPR_ACCOUNTREPORT_BULANAN o " & _ 
"		where to_char(o.transdate,'mm')<='" & bulansplitfrom & "' and to_char(o.transdate,'yyyy')='" & thnsplitfrom & "' AND o.source='3.VRA' " & _ 
"		GROUP BY o.fcba,o.fccode,o.fcname,o.uom_unit,o.sectioncode,o.fieldcode,o.fieldarea,o.fieldstatus, o.STOCKCODE || ' - ' || o.STOCKNAME || ' - ' || o.stockuom,o.vracode || ' - ' || o.vraname " & _
"     )t " & _
	    filterlhd & _  
 	    filterdiv & _ 
		filterfccode & _
		filterBlok & _
"  GROUP BY t.fcba,t.sectioncode,t.fieldcode,t.fieldarea,t.fieldstatus,t.fccode,t.fcname,t.itembahan,t.itemkendaraan,t.uom_unit " & _
" ORDER BY t.fcba,t.fccode,t.fcname,t.sectioncode,t.fieldcode,t.fieldarea,t.fieldstatus,t.itembahan"

' response.Write(sqlview1)

	Set rsCat = Server.CreateObject("ADODB.Recordset")
			rsCat.CursorType = 3
			rsCat.CursorLocation = 3
			rsCat.Open sqlview1, connstr


%>

<body>
	<div id="wrapper">
        <!--#include file="print_excel_blob.asp"-->
		<table border="0" width="100%" id="product-table" class="xthf-site">
			<thead>
				<tr>
					
					<th colspan="48" style=" background-color:#FFFFFF; color:#000000; border:none">
						<div id="header-wrapper">
							<div id="head_judul" style="float:left; text-align:left;">
								<h4>PT. <%=CompanyName%></h4>
								<h5>Estate: <%=fcba%></h5>
							</div>
							
						</div>
					</th>
				</tr>
						
				<tr>
						
					<th align="left" colspan="48" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">
						LAPORAN HARIAN DIVISI REKAP
					</th>
						
				</tr>
				<tr>
						
					<th align="left" colspan="48" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:center;">
						Tanggal Awal: 
						<%=tglfrom%>
						s/d
						<%=tglTo%>
					</th>
				</tr>
						


	<%	
	
	NO=1%>
	<tr>
			<th rowspan="2">No</th>
			<th rowspan="2">Status</th>
			<th rowspan="2">Activity</th>
			<th rowspan="2">Blok</th>
			<th colspan="8">UPAH</th>
			<th colspan="7">BORONGAN</th>
			<th colspan="5">MATERIAL</th>
			<th colspan="4">KENDARAAN</th>
			<th colspan="2" rowspan="2">TOTAL</th>
	</tr>
	<tr>
			<th colspan="2">Ha</th>			
			<th colspan="2">Unit</th>
			<th colspan="2">HK</th>
			<th colspan="2">Rp.HK</th>
			<th colspan="2">Ha</th>
			<th colspan="2">Unit</th>
			<th colspan="2">Rp</th>
			<th rowspan="2">Item Bahan</th>
			<th colspan="2">Qty</th>
			<th colspan="2">Biaya</th>
			<th rowspan="2">Kode</th>
			<th colspan="2">HM/KM</th>
			<th colspan="2">Biaya</th>
		</tr>
</thead>
<tbody>
	<%
	
	   if Not rsCat.eof then
		
			no=1
		
		Dim ttlhk,ttlhkshi,ttlhksbi,ttlfisik,ttlunit, ttlrphkhi,ttlrphkshi,ttlrphksbi
		dim ttlfisikshi,ttlfisiksbi,ttlunitshi,ttlunitsbi
		dim ttlfisikbrhi,ttlfisikbrshi,ttlfisikbrsbi,ttlunitbrhi,ttlunitbrshi,ttlunitbrsbi
		dim ttlrpbrhi,ttlrpbrshi,ttlrpbrsbi
		dim ttlqtybahanhi,ttlqtybahanshi,ttlqtybahansbi,ttlbiayabahanhi,ttlbiayabahanshi,ttlbiayabahansbi
		dim ttlqtyvrahi,ttlqtyvrashi,ttlqtyvrasbi,ttlbiayavrahi,ttlbiayavrashi,ttlbiayavrasbi
		dim ttlrupiahhi,ttlrupiahshi,ttlrupiahsbi
		dim ttlgrndrupiahhi,ttlgrndrupiahshi,ttlgrndrupiahsbi

			ttlfisik = 0
			ttlfisikshi = 0
			ttlfisiksbi = 0
			ttlunit = 0
			ttlunitshi = 0
			ttlunitsbi = 0
			ttlhk = 0
			ttlhkshi = 0
			ttlhksbi = 0
			
			ttlrphkhi = 0
			ttlrphkshi = 0
			ttlrphksbi = 0

			ttlfisikbrhi = 0
			ttlfisikbrshi = 0
			ttlfisikbrsbi = 0

			ttlunitbrhi = 0
			ttlunitbrshi = 0
			ttlunitbrsbi = 0

			ttlrpbrhi = 0
			ttlrpbrshi = 0
			ttlrpbrsbi = 0

			ttlqtybahanhi = 0
			ttlqtybahanshi = 0
			ttlqtybahansbi = 0
			ttlbiayabahanhi = 0
			ttlbiayabahanshi = 0
			ttlbiayabahansbi = 0
			 ttlqtyvrahi	= 0
			 ttlqtyvrashi = 0
			 ttlqtyvrasbi = 0
			 ttlbiayavrahi = 0
			 ttlbiayavrashi = 0
			 ttlbiayavrasbi = 0
			 ttlgrndrupiahhi = 0
			 ttlgrndrupiahshi = 0
			 ttlgrndrupiahsbi = 0

		While Not rsCat.EOF
			ttlrupiahhi = cdbl(rscat("rphkhi"))+cdbl(rscat("rpbrhi"))+cdbl(rscat("biayabahanhi"))+cdbl(rscat("biayavrahi"))
			ttlrupiahshi = cdbl(rscat("rphkshi"))+cdbl(rscat("rpbrshi"))+cdbl(rscat("biayabahanshi"))+cdbl(rscat("biayavrashi"))
			ttlrupiahsbi = cdbl(rscat("rphksbi"))+cdbl(rscat("rpbrsbi"))+cdbl(rscat("biayabahansbi"))+cdbl(rscat("biayavrasbi"))
			%>
		<tr>
			
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=no%></td>		
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fieldstatus")%></td>
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fccode") & "-" & rscat("fcname")%></td>					
			<td align="center" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=rscat("fieldcode")%></td>
			<td align="Right" colspan="2" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("fisikshi"),2)%></td>				
			<td align="Right" colspan="2" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("unitshi"),2)%></td>	
			
			<td align="Right" colspan="2" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("hkshi"),2)%></td>	
			<td align="right" colspan="2" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("rphkshi"),2)%></td>
				

			<td align="right" colspan="2" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("fisikbrshi"),2)%></td>	
			<td align="right" colspan="2" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("unitbrshi"),2)%></td>	
			
			<td align="right" colspan="2" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("rpbrshi"),2)%></td>		
	
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=(rscat("itembahan"))%></td>	
			<td align="right" colspan="2" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("qtybahanshi"),2)%></td>	
			<td align="right" colspan="2" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("biayabahanshi"),2)%></td>		
	
			<td align="left" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=(rscat("itemkendaraan"))%></td>	
			<td align="right" colspan="2" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("qtyvrashi"),2)%></td>	
			<td align="right" colspan="2" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(rscat("biayavrashi"))%></td>	
			<td align="right" colspan="2" style="<%=cmdstyle%>"><font face="Verdana" size="1"><%=FormatNumber(ttlrupiahshi)%></td>	
		</tr>
	  <%
	  
		ttlfisik = cdbl(ttlfisik)+cdbl(rscat("fisik"))
		ttlfisikshi = cdbl(ttlfisikshi)+cdbl(rscat("fisikshi"))
		ttlfisiksbi = cdbl(ttlfisiksbi)+cdbl(rscat("fisiksbi"))

		ttlunit = cdbl(ttlunit)+cdbl(rscat("unit"))
		ttlunitshi = cdbl(ttlunit)+cdbl(rscat("unitshi"))
		ttlunitsbi = cdbl(ttlunit)+cdbl(rscat("unitsbi"))

		ttlhk = cdbl(ttlhk)+cdbl(rscat("hk"))
		ttlhkshi = cdbl(ttlhkshi)+cdbl(rscat("hkshi"))
		ttlhksbi = cdbl(ttlhksbi)+cdbl(rscat("hksbi"))
		ttlrphkhi = cdbl(ttlrphkhi)+cdbl(rscat("rphkhi"))
		ttlrphkshi = cdbl(ttlrphkshi)+cdbl(rscat("rphkshi"))
		ttlrphksbi = cdbl(ttlrphksbi)+cdbl(rscat("rphksbi"))

		ttlfisikbrhi = cdbl(ttlfisikbrhi)+cdbl(rscat("fisikbrhi"))
		ttlfisikbrshi = cdbl(ttlfisikbrshi)+cdbl(rscat("fisikbrshi"))
		ttlfisikbrsbi = cdbl(ttlfisikbrsbi)+cdbl(rscat("fisikbrsbi"))

		ttlunitbrhi = cdbl(ttlunitbrhi)+cdbl(rscat("unitbrhi"))
		ttlunitbrshi = cdbl(ttlunitbrshi)+cdbl(rscat("unitbrshi"))
		ttlunitbrsbi = cdbl(ttlunitbrsbi)+cdbl(rscat("unitbrsbi"))

		ttlrpbrhi = cdbl(ttlrpbrhi)+cdbl(rscat("rpbrhi"))
		ttlrpbrshi = cdbl(ttlrpbrshi)+cdbl(rscat("rpbrshi"))
		ttlrpbrsbi = cdbl(ttlrpbrsbi)+cdbl(rscat("rpbrsbi"))

		ttlqtybahanhi = cdbl(ttlqtybahanhi) + cdbl(rscat("qtybahanhi"))
		ttlqtybahanshi = cdbl(ttlqtybahanshi) + cdbl(rscat("qtybahanshi"))
		ttlqtybahansbi = cdbl(ttlqtybahansbi) + cdbl(rscat("qtybahansbi"))
		ttlbiayabahanhi = cdbl(ttlbiayabahanhi) + cdbl(rscat("biayabahanhi"))
		ttlbiayabahanshi = cdbl(ttlbiayabahanshi) + cdbl(rscat("biayabahanshi"))
		ttlbiayabahansbi = cdbl(ttlbiayabahansbi) + cdbl(rscat("biayabahansbi"))

		ttlqtyvrahi	= cdbl(ttlqtyvrahi) + cdbl(rscat("qtyvrahi"))
		ttlqtyvrashi = cdbl(ttlqtyvrashi) + cdbl(rscat("qtyvrashi"))
		ttlqtyvrasbi = cdbl(ttlqtyvrasbi) + cdbl(rscat("qtyvrasbi"))
		ttlbiayavrahi = cdbl(ttlbiayavrahi) + cdbl(rscat("biayavrahi"))
		ttlbiayavrashi = cdbl(ttlbiayavrashi) + cdbl(rscat("biayavrashi"))
		ttlbiayavrasbi = cdbl(ttlbiayavrasbi) + cdbl(rscat("biayavrasbi"))

		ttlgrndrupiahhi = cdbl(ttlgrndrupiahhi) + cdbl(ttlrupiahhi)
		ttlgrndrupiahshi = cdbl(ttlgrndrupiahshi) + cdbl(ttlrupiahshi)
		ttlgrndrupiahsbi = cdbl(ttlgrndrupiahsbi) + cdbl(ttlrupiahsbi)

	  rsCat.MoveNext

	  no=no+1
	Wend
	rsCat.close()
	rsCat=null
	END IF
%>

		<tr bgcolor="#F4F4A0">
			<td align="Center" style="<%=cmdstyle%>" colspan= "4"><font face="Verdana" size="1"><strong>T o t a l</strong></td>
			<td align="Right" colspan="2" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlfisikshi, 2)%></strong></td>
			<td align="Right" colspan="2" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlunitshi, 2)%></strong></td>
			<td align="Right" colspan="2" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlhkshi, 2)%></strong></td> 
			<td align="Right" colspan="2" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlrphkshi, 2)%></strong></td>
			<td align="Right" colspan="2" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlfisikbrshi, 2)%></strong></td>
			<td align="Right" colspan="2" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlunitbrshi, 2)%></strong></td>
			<td align="Right" colspan="2"style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlrpbrshi, 2)%></strong></td>
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"></td>
			<td align="Right" colspan="2" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlqtybahanshi, 2)%></strong></td>
			<td align="Right" colspan="2" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlbiayabahanshi, 2)%></strong></td>			
			<td align="Right" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong></strong></td>
			<td align="Right" colspan="2" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlqtyvrashi, 2)%></strong></td>
			<td align="Right" colspan="2" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlbiayavrashi, 2)%></strong></td>
			<td align="Right" colspan="2" style="<%=cmdstyle%>"> <font face="Verdana" size="1"><strong><%=FormatNumber(ttlgrndrupiahshi, 2)%></strong></td>
		</tr>
	</tbody>
</table>
<!--#include file="Bottom_Report.asp"---------->

		</DIV>
	</BODY>
</HTML>	