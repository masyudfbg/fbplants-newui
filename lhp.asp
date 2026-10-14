<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<html>
	<head>
		<title>ePlaS - Laporan Harian Produksi</title>
		<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<link rel="StyleSheet" href="style-content.css" type="text/css" />
<link rel="stylesheet" href="css/jquery-ui.css" />
<script type="text/javascript" src="dtree.js"></script>
<script src="/js/jquery-1.9.1.js"></script>
<script src="/js/jquery-ui.js"></script>
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

<script language="JavaScript">

var newwin;
function launchwin(winurl,winname,winfeatures)
{
	//This launches a new window and then
	//focuses it if window.focus() is supported.
	newwin = window.open(winurl,winname,winfeatures);

}

function harusangka(evt) {
    var e = evt || window.event;
    var key = e.keyCode || e.which;

    if (!e.shiftKey && !e.altKey && !e.ctrlKey &&
    // numbers   
    key >= 48 && key <= 57 ||
	e.ctrlKey && key == 97 ||
    // Backspace and Tab and Enter
    key == 8 || key == 9 || key == 13 ||
    // left and right arrows
    key == 37 || key == 39 ||
    // Del and Ins
	key == 46 ) {
        // input is VALID
		}
	else {
			// input is INVALID
			e.returnValue = false;
			if (e.preventDefault) e.preventDefault();
		}
}


function getTotal()
{
	var tk1 = $("#isicpo1").val();
	var tk2 = $("#isicpo2").val();
	var tk3 = $("#isicpo3").val();
	var total = (parseInt(tk1, 10) || 0) + (parseInt(tk2, 10) || 0) + (parseInt(tk3, 10) || 0);   // kolom kosong dihitung 0 (dulu jadi NaN)
	$("#totalcpo").val(total);
}
</script>
<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css + css/fbp-docdetail.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="css/fbp-forms.css?v=20261009e" type="text/css" />
<link rel="StyleSheet" href="css/fbp-docdetail.css?v=20261009b" type="text/css" />
<script type="text/javascript" src="js/fbp-shell.js?v=20261014b"></script>
<% end if %>
</head>
<style>
	.posisikirijudul{
		position: relative;
		left: 0px;
	}
	.posisikiriisi{
		position: relative;
		left: 0px;
	}
	.lebarkolomsatu{
		width: 10px;
	}
	.lebarkolomdua{
		width: 10px;
	}
	.warnagaris{
		border-color: #FFA800;
	}
	.tebal{
		font-weight: bold;
	}
	.ratakanan{
		text-align: right;
	}
	.ratatengah{
		text-align: center;
	}
	/* ---- Tampilan baru Laporan Harian Produksi (hanya tampilan; nama field & proses simpan tetap) ---- */
	.lhp { --brand: #1B7863; --brand-dark: #12584a; --tint: #eef6f3; --line: #e1e7e5; --muted: #5d6d69; --ink: #17262a; color: var(--ink); padding: 2px 2px 8px; }
	.lhp-head { margin-bottom: 14px; }
	.lhp-eyebrow { display: inline-flex; align-items: center; gap: 6px; font-size: 12px; font-weight: 600; letter-spacing: .3px; text-transform: uppercase; color: var(--brand-dark); }
	.lhp-head h2 { margin: 4px 0 2px; font-size: 21px; font-weight: 700; color: var(--ink); }
	.lhp-head p { margin: 0; font-size: 13px; color: var(--muted); }
	.lhp-kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 10px; margin-bottom: 16px; }
	.lhp-kpi { position: relative; padding: 12px 15px 12px 16px; background: #fff; border: 1px solid var(--line); border-radius: 12px; overflow: hidden; }
	.lhp-kpi::before { content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 4px; background: var(--brand); opacity: .75; }
	.lhp-kpi span { display: flex; align-items: center; gap: 6px; font-size: 11.5px; font-weight: 600; letter-spacing: .3px; text-transform: uppercase; color: var(--muted); }
	.lhp-kpi span .fa { color: var(--brand); }
	.lhp-kpi b { display: inline-block; margin-top: 5px; font-size: 20px; font-weight: 700; font-variant-numeric: tabular-nums; }
	.lhp-kpi small { margin-left: 5px; font-size: 12px; color: var(--muted); }
	.lhp-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; }
	.lhp-card { min-width: 0; padding: 16px 18px 18px; background: #fff; border: 1px solid var(--line); border-radius: 14px; }
	.lhp-card.wide { grid-column: 1 / -1; }
	.lhp-card h3 { display: flex; align-items: center; gap: 10px; margin: 0 0 14px; font-size: 15px; font-weight: 700; color: var(--ink); }
	.lhp-card h3 .ic { display: inline-grid; place-items: center; width: 30px; height: 30px; font-size: 14px; color: var(--brand-dark); background: var(--tint); border-radius: 8px; }
	.lhp-card h3 small { font-size: 12px; font-weight: 500; color: var(--muted); }
	.lhp-fields { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px 14px; }
	.lhp-fields.cols4 { grid-template-columns: repeat(4, minmax(0, 1fr)); }
	.lhp-f { display: flex; flex-direction: column; gap: 5px; min-width: 0; }
	.lhp-f .lb { font-size: 12.5px; font-weight: 600; color: #3d4d49; }
	.lhp-f .ig { position: relative; display: flex; }
	.box_both .lhp .lhp-f .ig input { width: 100% !important; min-width: 0; height: 40px !important; margin: 0 !important; padding: 0 12px !important; font-size: 14px; font-variant-numeric: tabular-nums; border: 1px solid #d5dfdc !important; border-radius: 9px !important; box-sizing: border-box; }
	.box_both .lhp .lhp-f .ig input:not(:last-child) { border-top-right-radius: 0 !important; border-bottom-right-radius: 0 !important; }
	.lhp-f .ig em { flex: none; display: grid; place-items: center; min-width: 46px; padding: 0 10px; font-size: 12px; font-style: normal; font-weight: 600; color: var(--muted); background: #f3f6f5; border: 1px solid #d5dfdc; border-left: 0; border-radius: 0 9px 9px 0; box-sizing: border-box; }
	.box_both .lhp input[readonly] { color: #3d4d49 !important; background: #f2f5f4 !important; }
	.box_both .lhp input:not([readonly]):focus, .box_both .lhp textarea:focus { border-color: var(--brand) !important; box-shadow: 0 0 0 3px rgba(27,120,99,.16) !important; outline: none; }
	.lhp-sub { margin: 18px 0 8px; font-size: 12px; font-weight: 700; letter-spacing: .3px; text-transform: uppercase; color: var(--muted); }
	.lhp-tblwrap { overflow-x: auto; border: 1px solid var(--line); border-radius: 10px; }
	.box_both .lhp table.lhp-tbl { width: 100%; margin: 0; border-collapse: collapse; border: 0; }
	.box_both .lhp table.lhp-tbl th { padding: 10px 12px !important; font-size: 11.5px; font-weight: 700; letter-spacing: .3px; text-transform: uppercase; text-align: center; color: var(--muted) !important; background: #f6f9f8 !important; border: 0 !important; border-bottom: 1px solid var(--line) !important; white-space: nowrap; }
	.box_both .lhp table.lhp-tbl td { padding: 8px 10px !important; text-align: center; border: 0 !important; border-bottom: 1px solid #edf1f0 !important; background: #fff; }
	.box_both .lhp table.lhp-tbl tbody tr:last-child td { border-bottom: 0 !important; }
	.box_both .lhp table.lhp-tbl tfoot td { font-weight: 700; background: #f6f9f8; border-top: 1px solid var(--line) !important; }
	.box_both .lhp table.lhp-tbl input { width: 100% !important; min-width: 80px; height: 38px !important; margin: 0 !important; padding: 0 10px !important; font-size: 14px; font-variant-numeric: tabular-nums; border: 1px solid #d5dfdc !important; border-radius: 8px !important; box-sizing: border-box; }
	.box_both .lhp table.lhp-tbl input.tank { width: 56px !important; min-width: 56px; text-align: center !important; font-weight: 700; color: var(--brand-dark) !important; background: var(--tint) !important; border-color: #cfe3dc !important; }
	.box_both .lhp textarea { width: 100% !important; min-height: 84px; padding: 10px 12px !important; font: 14px "Plus Jakarta Sans", system-ui, sans-serif; border: 1px solid #d5dfdc !important; border-radius: 9px !important; box-sizing: border-box; resize: vertical; }
	.lhp-bar { position: sticky; bottom: 0; z-index: 5; display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 10px; margin-top: 14px; padding: 12px 16px; background: rgba(255,255,255,.96); border: 1px solid var(--line); border-radius: 12px; box-shadow: 0 -6px 18px rgba(23,38,42,.06); font-size: 13px; color: var(--muted); }
	.box_both .lhp .lhp-save { display: inline-flex; align-items: center; gap: 8px; min-height: 42px; padding: 0 26px !important; font-size: 14px; font-weight: 700; color: #fff !important; background: var(--brand) !important; border: 1px solid var(--brand) !important; border-radius: 10px !important; cursor: pointer; }
	.box_both .lhp .lhp-save:hover { background: var(--brand-dark) !important; }
	@media (max-width: 1100px) { .lhp-fields.cols4 { grid-template-columns: repeat(2, minmax(0, 1fr)); } }
	@media (max-width: 900px) { .lhp-grid { grid-template-columns: 1fr; } }
	@media (max-width: 520px) { .lhp-fields, .lhp-fields.cols4 { grid-template-columns: 1fr; } }
</style>
<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy page-doc"<% end if %>>
<!--#include file="top.jsp"-->
<%
Dim sqltgl, sqltbs,rsdata,sqlsimpan,sqlsimpanhk,sqlbulking,sqltankcpo
Dim nettotbs,nettocpo,nettokernel,nettockg,fcba
Dim hk1,cuti1,sakit1,izin1,mangkir1
Dim hk2,cuti2,sakit2,izin2,mangkir2
Dim txttgl,jamkerja,terimatbs,sisatbs
Dim kirimcpo,cotcpo,cstcpo,otcpo,ffacpo,aircpo,kotorancpo,dobicpo
Dim kirimkernel,silokernel,kirimckg,produksickg
Dim bulkingsilo,isibulking,ffabulking,airbulking,kotoranbulking
Dim TankID1,isicpo1,ffacpo1,aircpo1,kotorancpo1
Dim TankID2,isicpo2,ffacpo2,aircpo2,kotorancpo2
Dim TankID3,isicpo3,ffacpo3,aircpo3,kotorancpo3
Dim totalcpo,air,caco3,keterangan,jammulai,jamselesai,sqldelete,pottbskeluar,adjcpo

	

%>
<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
			<!--#include file="menu.asp"-->
</div>
<div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
<div class="judul">
	   Module : Produksi Harian
</div>
<div class="box_both">
<%

		fcba = Session("fcba")
		
		' GET PRODUCTION DATE'
		sqltgl = "select to_char(PRODUCTIONMILLDATE,'mm/dd/yyyy') PRODUCTIONMILLDATE  " & _
		" from CONFIGMST " & _
		" where fcba='" & fcba & "'"
		
		Set rsdata = Server.CreateObject("ADODB.Recordset")
		rsdata.CursorType = 3
		rsdata.CursorLocation = 3
		rsdata.Open sqltgl, connstr
		if not rsdata.eof then
			txttgl = rsdata("PRODUCTIONMILLDATE")
		end if
		rsdata.Close
		set rsdata = Nothing
		
		' GET NETTO TBS'
		sqltbs = "select NVL(SUM(WEIGHTIN-WEIGHTOUT),0) AS NETTOTBS " & _
		" from LHPTIKETHDR " & _
		" where DOCKETDATE= TO_DATE('" & txttgl & "', 'mm/dd/yyyy') AND fcba='" & fcba & "' " & _ 
		" AND PRODUCTID='2'"
		'response.write sqltbs
		Set rsdata = Server.CreateObject("ADODB.Recordset")
		rsdata.CursorType = 3
		rsdata.CursorLocation = 3
		rsdata.Open sqltbs, connstr
		if not rsdata.eof then
			nettotbs = cdbl(to_nol(rsdata("NETTOTBS")))
		end if
		rsdata.Close
		set rsdata = Nothing

		' GET NETTO CPO'
		sqltbs = "select NVL(SUM(WEIGHTOUT-WEIGHTIN),0) AS NETTOCPO " & _
		" from LHPTIKETHDR h,LHPTIKETCPO d,LHPDOMST dot,LHPCONTRACTMST c " & _
		" where h.docketid=d.docketid AND d.doid=dot.doid AND dot.contractid=c.contractid AND h.DOCKETDATE= TO_DATE('" & txttgl & "', 'mm/dd/yyyy') AND h.fcba='" & fcba & "' " & _ 
		" AND h.PRODUCTID='1' and c.tipekontrakid IN('1')"
		'response.write sqltbs
		Set rsdata = Server.CreateObject("ADODB.Recordset")
		rsdata.CursorType = 3
		rsdata.CursorLocation = 3
		rsdata.Open sqltbs, connstr
		if not rsdata.eof then
			nettocpo = cdbl(to_nol(rsdata("NETTOCPO")))
		end if
		rsdata.Close
		set rsdata = Nothing


		' GET NETTO KERNEL'
		sqltbs = "select NVL(SUM(WEIGHTOUT-WEIGHTIN),0) AS NETTOKERNEL " & _
		" from LHPTIKETHDR " & _
		" where DOCKETDATE= TO_DATE('" & txttgl & "', 'mm/dd/yyyy') AND fcba='" & fcba & "' " & _ 
		" AND PRODUCTID='3'  AND ISINSIDEFLG=0"
		'response.write sqltbs
		Set rsdata = Server.CreateObject("ADODB.Recordset")
		rsdata.CursorType = 3
		rsdata.CursorLocation = 3
		rsdata.Open sqltbs, connstr
		if not rsdata.eof then
			nettokernel = cdbl(to_nol(rsdata("NETTOKERNEL")))
		end if
		rsdata.Close
		set rsdata = Nothing

		' GET NETTO CKG'
		sqltbs = "select NVL(SUM(WEIGHTOUT-WEIGHTIN),0) AS NETTOCKG " & _
		" from LHPTIKETHDR " & _
		" where DOCKETDATE= TO_DATE('" & txttgl & "', 'mm/dd/yyyy') AND fcba='" & fcba & "' " & _ 
		" AND PRODUCTID='5'"
		'response.write sqltbs
		Set rsdata = Server.CreateObject("ADODB.Recordset")
		rsdata.CursorType = 3
		rsdata.CursorLocation = 3
		rsdata.Open sqltbs, connstr
		if not rsdata.eof then
			nettockg = cdbl(to_nol(rsdata("NETTOCKG")))
		end if
		rsdata.Close
		set rsdata = Nothing

		' GET HK PROSES1'
		sqltbs = "select gangcode,sum(hk1) as hk1,sum(cuti1) as cuti1,sum(sakit1) as sakit1,sum(izin1) as izin1,sum(mangkir1) as mangkir1 " & _
		" from( " & _
		"	select gangcode,count(mandays) as hk1,0 AS cuti1,0 as sakit1,0 as izin1,0 as mangkir1 " & _
		"	from ATTENDANCE_GAD_TEMP " & _
		"	WHERE FDDATE=TO_DATE('" & txttgl & "','mm/dd/yyyy') and attendance = 'KJ' and gangcode='PR001' " & _
		"	group by gangcode " & _
		"		union all " & _
		"	select gangcode,0 as hk1,count(mandays) as cuti1,0 as sakit1,0 as izin1,0 as mangkir1 " & _
		"	from ATTENDANCE_GAD_TEMP " & _
		"	WHERE FDDATE=TO_DATE('" & txttgl & "','mm/dd/yyyy') and attendance = 'CT' and gangcode='PR001' " & _
		"	group by gangcode  " & _
		"		union all " & _
		"	select gangcode,0 as hk1,0 as cuti1,count(mandays) as sakit1,0 as izin1,0 as mangkir1 " & _
		"	from ATTENDANCE_GAD_TEMP " & _
		"	WHERE FDDATE=TO_DATE('" & txttgl & "','mm/dd/yyyy') and attendance = 'S1' and gangcode='PR001' " & _
		"	group by gangcode  " & _
		"		union all " & _
		"	select gangcode,0 as hk1,0 as cuti1,0 as sakit1,count(mandays) as izin1,0 as mangkir1 " & _
		"	from ATTENDANCE_GAD_TEMP " & _
		"	WHERE FDDATE=TO_DATE('" & txttgl & "','mm/dd/yyyy') and attendance = 'P1' and gangcode='PR001' " & _
		"	group by gangcode  " & _
		"		union all " & _
		"	select gangcode,0 as hk1,0 as cuti1,0 as sakit1,0 as izin1,count(mandays) as mangkir1 " & _
		"	from ATTENDANCE_GAD_TEMP " & _
		"	WHERE FDDATE=TO_DATE('" & txttgl & "','mm/dd/yyyy') and attendance = 'MK' and gangcode='PR001' " & _
		"	group by gangcode " & _
		" ) " & _
		" group by gangcode"
		'response.write sqltbs
		Set rsdata = Server.CreateObject("ADODB.Recordset")
		rsdata.CursorType = 3
		rsdata.CursorLocation = 3
		rsdata.Open sqltbs, connstr

		if not rsdata.eof then
			hk1 		= rsdata("hk1")
			cuti1 		= rsdata("cuti1")
			sakit1 		= rsdata("sakit1")
			izin1 		= rsdata("izin1")
			mangkir1 	= rsdata("mangkir1")
		end if
		rsdata.Close
		set rsdata = Nothing

		' GET HK PROSES2'
		sqltbs = "select gangcode,sum(hk2) as hk2,sum(cuti2) as cuti2,sum(sakit2) as sakit2,sum(izin2) as izin2,sum(mangkir2) as mangkir2 " & _
		" from( " & _
		"	select gangcode,count(mandays) as hk2,0 AS cuti2,0 as sakit2,0 as izin2,0 as mangkir2 " & _
		"	from ATTENDANCE_GAD_TEMP " & _
		"	WHERE FDDATE=TO_DATE('" & txttgl & "','mm/dd/yyyy') and attendance = 'KJ' and gangcode='PR002' " & _
		"	group by gangcode " & _
		"		union all " & _
		"	select gangcode,0 as hk2,count(mandays) as cuti2,0 as sakit2,0 as izin2,0 as mangkir2 " & _
		"	from ATTENDANCE_GAD_TEMP " & _
		"	WHERE FDDATE=TO_DATE('" & txttgl & "','mm/dd/yyyy') and attendance = 'CT' and gangcode='PR002' " & _
		"	group by gangcode  " & _
		"		union all " & _
		"	select gangcode,0 as hk2,0 as cuti2,count(mandays) as sakit2,0 as izin2,0 as mangkir2 " & _
		"	from ATTENDANCE_GAD_TEMP " & _
		"	WHERE FDDATE=TO_DATE('" & txttgl & "','mm/dd/yyyy') and attendance = 'S1' and gangcode='PR002' " & _
		"	group by gangcode  " & _
		"		union all " & _
		"	select gangcode,0 as hk2,0 as cuti2,0 as sakit2,count(mandays) as izin2,0 as mangkir2 " & _
		"	from ATTENDANCE_GAD_TEMP " & _
		"	WHERE FDDATE=TO_DATE('" & txttgl & "','mm/dd/yyyy') and attendance = 'P1' and gangcode='PR002' " & _
		"	group by gangcode  " & _
		"		union all " & _
		"	select gangcode,0 as hk2,0 as cuti2,0 as sakit2,0 as izin2,count(mandays) as mangkir2 " & _
		"	from ATTENDANCE_GAD_TEMP " & _
		"	WHERE FDDATE=TO_DATE('" & txttgl & "','mm/dd/yyyy') and attendance = 'MK' and gangcode='PR002' " & _
		"	group by gangcode " & _
		" ) " & _
		" group by gangcode"
		'response.write sqltbs
		Set rsdata = Server.CreateObject("ADODB.Recordset")
		rsdata.CursorType = 3
		rsdata.CursorLocation = 3
		rsdata.Open sqltbs, connstr
		if not rsdata.eof then
			hk2 		= rsdata("hk2")
			cuti2 		= rsdata("cuti2")
			sakit2 		= rsdata("sakit2")
			izin2 		= rsdata("izin2")
			mangkir2 	= rsdata("mangkir2")
		end if
		rsdata.Close
		set rsdata = Nothing
			
	%>
<%
	' Ringkasan untuk kartu atas (nilai input tetap angka mentah agar proses simpan tidak berubah)
	Dim lhpHk
	lhpHk = CDbl(to_nol(hk1)) + CDbl(to_nol(hk2))
%>
<div class="lhp">
	<div class="lhp-head">
		<div>
			<div class="lhp-eyebrow"><i class="fa fa-industry" aria-hidden="true"></i> Laporan Harian Produksi</div>
			<h2>Input Produksi Harian</h2>
			<p>Tanggal produksi <b><%=txttgl%></b>. Nilai berwarna abu-abu diambil otomatis dari timbangan &amp; absensi.</p>
		</div>
	</div>

	<div class="lhp-kpis">
		<div class="lhp-kpi"><span><i class="fa fa-truck" aria-hidden="true"></i> TBS Diterima</span><b><%=FormatNumber(to_nol(nettotbs), 0)%></b><small>KG</small></div>
		<div class="lhp-kpi"><span><i class="fa fa-tint" aria-hidden="true"></i> CPO Dikirim</span><b><%=FormatNumber(to_nol(nettocpo), 0)%></b><small>KG</small></div>
		<div class="lhp-kpi"><span><i class="fa fa-circle" aria-hidden="true"></i> Kernel Dikirim</span><b><%=FormatNumber(to_nol(nettokernel), 0)%></b><small>KG</small></div>
		<div class="lhp-kpi"><span><i class="fa fa-archive" aria-hidden="true"></i> Cangkang Dikirim</span><b><%=FormatNumber(to_nol(nettockg), 0)%></b><small>KG</small></div>
		<div class="lhp-kpi"><span><i class="fa fa-users" aria-hidden="true"></i> Hari Kerja</span><b><%=FormatNumber(lhpHk, 0)%></b><small>HK shift 1 + 2</small></div>
	</div>

<form name="frmEntry" method="post" action="lhpinsertproses.asp" class="lhp-form">
	<div class="lhp-grid">

		<!-- Produksi & TBS -->
		<section class="lhp-card">
			<h3><span class="ic"><i class="fa fa-calendar" aria-hidden="true"></i></span>Produksi &amp; TBS</h3>
			<div class="lhp-fields">
				<label class="lhp-f"><span class="lb">Tanggal</span>
					<span class="ig"><input type="text" maxlength="15" name="txttgl" readonly id="txttgl" size="15" value="<%=txttgl%>" title="Tanggal Produksi"></span></label>
				<label class="lhp-f"><span class="lb">Jam Kerja</span>
					<span class="ig"><input type="text" class="ratakanan" name="jamkerja" id="jamkerja" size="15" value="0" onkeypress="return harusangka(event);"><em>Jam</em></span></label>
				<label class="lhp-f"><span class="lb">Jam Mulai</span>
					<span class="ig"><input type="text" class="ratakanan" name="jammulai" id="jammulai" size="15" value="0" onkeypress="return harusangka(event);"></span></label>
				<label class="lhp-f"><span class="lb">Jam Selesai</span>
					<span class="ig"><input type="text" class="ratakanan" name="jamselesai" id="jamselesai" size="15" value="0" onkeypress="return harusangka(event);"></span></label>
				<label class="lhp-f"><span class="lb">TBS Terima</span>
					<span class="ig"><input type="text" class="ratakanan" name="terimatbs" id="terimatbs" size="15" readOnly value="<%=nettotbs%>"><em>KG</em></span></label>
				<label class="lhp-f"><span class="lb">TBS Sisa</span>
					<span class="ig"><input type="text" class="ratakanan" name="sisatbs" id="sisatbs" size="15" value="0" onkeypress="return harusangka(event);"><em>KG</em></span></label>
				<label class="lhp-f"><span class="lb">TBS Keluar</span>
					<span class="ig"><input type="text" class="ratakanan" name="pottbskeluar" id="pottbskeluar" size="15" value="0" onkeypress="return harusangka(event);"><em>KG</em></span></label>
			</div>
		</section>

		<!-- Cangkang, air & CaCO3 -->
		<section class="lhp-card">
			<h3><span class="ic"><i class="fa fa-archive" aria-hidden="true"></i></span>Cangkang, Air &amp; CaCO3</h3>
			<div class="lhp-fields">
				<label class="lhp-f"><span class="lb">Cangkang Kirim</span>
					<span class="ig"><input type="text" class="ratakanan" name="kirimckg" id="kirimckg" readOnly size="15" value="<%=nettockg%>"><em>KG</em></span></label>
				<label class="lhp-f"><span class="lb">Cangkang Produksi</span>
					<span class="ig"><input type="text" class="ratakanan" name="produksickg" id="produksickg" size="15" value="0" onkeypress="return harusangka(event);"><em>KG</em></span></label>
				<label class="lhp-f"><span class="lb">Pemakaian Air</span>
					<span class="ig"><input type="text" class="ratakanan" name="air" id="air" size="15" value="0" onkeypress="return harusangka(event);"><em>M3</em></span></label>
				<label class="lhp-f"><span class="lb">Pemakaian CaCO3</span>
					<span class="ig"><input type="text" class="ratakanan" name="caco3" id="caco3" size="15" value="0" onkeypress="return harusangka(event);"><em>KG</em></span></label>
			</div>
		</section>

		<!-- CPO -->
		<section class="lhp-card wide">
			<h3><span class="ic"><i class="fa fa-tint" aria-hidden="true"></i></span>CPO</h3>
			<div class="lhp-fields cols4">
				<label class="lhp-f"><span class="lb">Kirim</span>
					<span class="ig"><input type="text" class="ratakanan" name="kirimcpo" id="kirimcpo" readOnly size="15" value="<%=nettocpo%>"><em>KG</em></span></label>
				<label class="lhp-f"><span class="lb">Adjustment CPO</span>
					<span class="ig"><input type="text" class="ratakanan" name="adjcpo" id="adjcpo" size="15" value="0" onkeypress=""><em>KG</em></span></label>
				<label class="lhp-f"><span class="lb">C.O.T</span>
					<span class="ig"><input type="text" class="ratakanan" name="cotcpo" id="cotcpo" size="15" value="0" onkeypress="return harusangka(event);"><em>KG</em></span></label>
				<label class="lhp-f"><span class="lb">C.S.T</span>
					<span class="ig"><input type="text" class="ratakanan" name="cstcpo" id="cstcpo" size="15" value="0" onkeypress="return harusangka(event);"><em>KG</em></span></label>
				<label class="lhp-f"><span class="lb">O.T</span>
					<span class="ig"><input type="text" class="ratakanan" name="otcpo" id="otcpo" size="15" value="0" onkeypress="return harusangka(event);"><em>KG</em></span></label>
				<label class="lhp-f"><span class="lb">FFA</span>
					<span class="ig"><input type="text" class="ratakanan" name="ffacpo" id="ffacpo" size="15" value="0" onkeypress="return harusangka(event);"><em>%</em></span></label>
				<label class="lhp-f"><span class="lb">Air</span>
					<span class="ig"><input type="text" class="ratakanan" name="aircpo" id="aircpo" size="15" value="0" onkeypress="return harusangka(event);"><em>%</em></span></label>
				<label class="lhp-f"><span class="lb">Kotoran</span>
					<span class="ig"><input type="text" class="ratakanan" name="kotorancpo" id="kotorancpo" size="15" value="0" onkeypress="return harusangka(event);"><em>%</em></span></label>
				<label class="lhp-f"><span class="lb">DOBI</span>
					<span class="ig"><input type="text" class="ratakanan" name="dobicpo" id="dobicpo" size="15" value="0" onkeypress="return harusangka(event);"></span></label>
			</div>
			<div class="lhp-sub">Stok CPO per tangki</div>
			<div class="lhp-tblwrap">
				<table class="lhp-tbl" id="table1">
					<thead><tr><th>Tangki</th><th>Isi (Kg)</th><th>FFA (%)</th><th>Air (%)</th><th>Kotoran (%)</th></tr></thead>
					<tbody>
						<tr>
							<td><input type="text" size="12" class="tank" name="TankID1" value="1" readonly></td>
							<td><input type="text" size="12" class="ratakanan" name="isicpo1" id="isicpo1" value="0" onkeyup="getTotal()" onkeypress="return harusangka(event);"></td>
							<td><input type="text" size="12" class="ratakanan" name="ffacpo1" value="0" onkeypress="return harusangka(event);"></td>
							<td><input type="text" size="12" class="ratakanan" name="aircpo1" value="0" onkeypress="return harusangka(event);"></td>
							<td><input type="text" size="12" class="ratakanan" name="kotorancpo1" value="0" onkeypress="return harusangka(event);"></td>
						</tr>
						<tr>
							<td><input type="text" size="12" class="tank" name="TankID2" value="2" readonly></td>
							<td><input type="text" size="12" class="ratakanan" name="isicpo2" id="isicpo2" value="0" onkeyup="getTotal()" onkeypress="return harusangka(event);"></td>
							<td><input type="text" size="12" class="ratakanan" name="ffacpo2" value="0" onkeypress="return harusangka(event);"></td>
							<td><input type="text" size="12" class="ratakanan" name="aircpo2" value="0" onkeypress="return harusangka(event);"></td>
							<td><input type="text" size="12" class="ratakanan" name="kotorancpo2" value="0" onkeypress="return harusangka(event);"></td>
						</tr>
						<tr>
							<td><input type="text" size="12" class="tank" name="TankID3" value="3" readonly></td>
							<td><input type="text" size="12" class="ratakanan" name="isicpo3" id="isicpo3" value="0" onkeyup="getTotal()" onkeypress="return harusangka(event);"></td>
							<td><input type="text" size="12" class="ratakanan" name="ffacpo3" value="0" onkeypress="return harusangka(event);"></td>
							<td><input type="text" size="12" class="ratakanan" name="aircpo3" value="0" onkeypress="return harusangka(event);"></td>
							<td><input type="text" size="12" class="ratakanan" name="kotorancpo3" value="0" onkeypress="return harusangka(event);"></td>
						</tr>
					</tbody>
					<tfoot>
						<tr><td>Total</td><td><input type="text" size="12" class="ratakanan" name="totalcpo" readonly id="totalcpo" value="0"></td><td colspan="3"></td></tr>
					</tfoot>
				</table>
			</div>
		</section>

		<!-- Inti kelapa sawit -->
		<section class="lhp-card wide">
			<h3><span class="ic"><i class="fa fa-circle" aria-hidden="true"></i></span>Inti Kelapa Sawit (Kernel)</h3>
			<div class="lhp-fields cols4">
				<label class="lhp-f"><span class="lb">Kirim</span>
					<span class="ig"><input type="text" class="ratakanan" name="kirimkernel" id="kirimkernel" readOnly size="15" value="<%=NETTOKERNEL%>"><em>KG</em></span></label>
				<label class="lhp-f"><span class="lb">Kernel Silo</span>
					<span class="ig"><input type="text" class="ratakanan" name="silokernel" id="silokernel" size="15" value="0" onkeypress="return harusangka(event);"><em>KG</em></span></label>
			</div>
			<div class="lhp-sub">Bulking silo</div>
			<div class="lhp-tblwrap">
				<table class="lhp-tbl" id="table2">
					<thead><tr><th>Bulking Silo</th><th>Isi (Kg)</th><th>FFA (%)</th><th>Air (%)</th><th>Kotoran (%)</th></tr></thead>
					<tbody>
						<tr>
							<td><input type="text" size="12" class="tank" name="bulkingsilo" value="1" readonly onkeypress="return harusangka(event);"></td>
							<td><input type="text" size="12" class="ratakanan" name="isibulking" value="0" onkeypress="return harusangka(event);"></td>
							<td><input type="text" size="12" class="ratakanan" name="ffabulking" value="0" onkeypress="return harusangka(event);"></td>
							<td><input type="text" size="12" class="ratakanan" name="airbulking" value="0" onkeypress="return harusangka(event);"></td>
							<td><input type="text" size="12" class="ratakanan" name="kotoranbulking" value="0" onkeypress="return harusangka(event);"></td>
						</tr>
					</tbody>
				</table>
			</div>
		</section>

		<!-- Tenaga kerja -->
		<section class="lhp-card wide">
			<h3><span class="ic"><i class="fa fa-users" aria-hidden="true"></i></span>Tenaga Kerja <small>dari absensi (gang PR001 &amp; PR002)</small></h3>
			<div class="lhp-tblwrap">
				<table class="lhp-tbl" id="table3">
					<thead><tr><th>Shift</th><th>HK</th><th>Cuti</th><th>Sakit</th><th>Izin/P1</th><th>Mangkir</th></tr></thead>
					<tbody>
						<tr>
							<td><input type="text" class="tank" size="12" name="shift1" value="1" readonly></td>
							<td><input type="text" size="12" class="ratakanan" name="hk1" value="<%=hk1%>" readonly></td>
							<td><input type="text" size="12" class="ratakanan" name="cuti1" value="<%=cuti1%>" readonly></td>
							<td><input type="text" size="12" class="ratakanan" name="sakit1" value="<%=sakit1%>" readonly></td>
							<td><input type="text" size="12" class="ratakanan" name="izin1" value="<%=izin1%>" readonly></td>
							<td><input type="text" size="12" class="ratakanan" name="mangkir1" value="<%=mangkir1%>" readonly></td>
						</tr>
						<tr>
							<td><input type="text" class="tank" size="12" name="shift2" value="2" readonly></td>
							<td><input type="text" size="12" class="ratakanan" name="hk2" value="<%=hk2%>" readonly></td>
							<td><input type="text" size="12" class="ratakanan" name="cuti2" value="<%=cuti2%>" readonly></td>
							<td><input type="text" size="12" class="ratakanan" name="sakit2" value="<%=sakit2%>" readonly></td>
							<td><input type="text" size="12" class="ratakanan" name="izin2" value="<%=izin2%>" readonly></td>
							<td><input type="text" size="12" class="ratakanan" name="mangkir2" value="<%=mangkir2%>" readonly></td>
						</tr>
					</tbody>
				</table>
			</div>
		</section>

		<!-- Keterangan -->
		<section class="lhp-card wide">
			<h3><span class="ic"><i class="fa fa-pencil" aria-hidden="true"></i></span>Keterangan</h3>
			<textarea name="ket" cols="150" rows="3" placeholder="Isi keterangan (opsional)"></textarea>
		</section>
	</div>

	<div class="lhp-bar">
		<span><i class="fa fa-info-circle" aria-hidden="true"></i> Periksa kembali angka sebelum menyimpan.</span>
		<button type="submit" name="Add" value="Simpan" class="lhp-save"><i class="fa fa-check" aria-hidden="true"></i> Simpan</button>
	</div>
</form>
</div>
</div></div>
  <!--#include file="bottom.jsp"-->

	</body>
	<!-- xx-->
</html>
