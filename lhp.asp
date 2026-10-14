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
	var total = parseInt(tk1)+parseInt(tk2)+parseInt(tk3);
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
<script type="text/javascript" src="js/fbp-shell.js?v=20261014a"></script>
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
<div class="box_both" style="overflow:scroll;">
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
<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-file-text-o" aria-hidden="true"></i>Header Laporan Harian Produksi</div><% end if %>
<form name="frmEntry" method="post" action="lhpinsertproses.asp">
	
	<table class="data" border="0" width="1000px">
		<tr bgcolor="#00a300" style="color:#FFFFFF;" style="color:#FFFFFF;"">
			<td valign="top" align="center" colspan="4">PRODUKSI HARIAN</td>
		</tr>
			
		<tr>
			<td align="left" class="lebarkolomsatu">Tanggal</td>
			<td align="left" style="padding" class="lebarkolomdua">
				<input type="text" class="" maxlength="15" name="txttgl" readonly id="txttgl" size="15" value="<%=txttgl%>" title="Tanggal Produksi">
			</td>
			<td align="left" class="lebarkolomsatu">Jam Kerja</td>
			<td align="left" style="padding" class="lebarkolomdua">
				<input type="text" class="ratakanan"  name="jamkerja" id="jamkerja" size="15" value="0" onkeypress="return harusangka(event);"> Jam
			</td>
		</tr>
		<tr bgcolor="#00a300" style="color:#FFFFFF;" style="color:#FFFFFF;"">
			<td valign="top" align="center" colspan="4" >TBS</td>
		</tr>
		<tr>
			<td align="left" class="lebarkolomsatu">Terima</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" name="terimatbs" id="terimatbs" size="15" readOnly value="<%=nettotbs%>"> KG
			</td>
			<td align="left" class="lebarkolomsatu">Jam Mulai</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" name="jammulai" id="jammulai" size="15" value="0" onkeypress="return harusangka(event);"> 
			</td>
			
		</tr>
		<tr>
			<td align="left" class="lebarkolomsatu">Sisa</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" name="sisatbs" id="sisatbs" size="15" value="0" onkeypress="return harusangka(event);"> KG
			</td>
			<td align="left" class="lebarkolomsatu">Jam Selesai</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" name="jamselesai" id="jamselesai" size="15" value="0" onkeypress="return harusangka(event);">
			</td>
		</tr>
		<tr>
			<td align="left" class="lebarkolomsatu">TBS Keluar</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" name="pottbskeluar" id="pottbskeluar" size="15" value="0" onkeypress="return harusangka(event);"> KG
			</td>
			
		</tr>
	
		<tr bgcolor="#00a300" style="color:#FFFFFF;" style="color:#FFFFFF;"">
			<td valign="top" align="center" colspan="4" >CPO</td>
		</tr>
		<tr>
			<td align="left" class="lebarkolomsatu">Kirim</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" name="kirimcpo" id="kirimcpo" readOnly size="15" value="<%=nettocpo%>"> KG
			</td>
			<td colspan="2" rowspan="8">
				<table class="data warnagaris" align=left border="1"  width="100%" id='table1'>
					<tr >
						<td align=center class="tebal">Tanki</td>
						<td align=center class="tebal">Isi (Kg)</td>
						<td align=center class="tebal">FFA (%)</td>
						<td align=center class="tebal">Air (%)</td>
						<td align=center class="tebal">Kotoran (%)</td>
					</tr>
					<tr>
						<td align=center><INPUT type="text" size=12 name='TankID1' value='1' readonly></td>
						<td align=center><INPUT type="text" size=12 class='ratakanan' name='isicpo1' id='isicpo1' value='0' onkeyup="getTotal()" onkeypress="return harusangka(event);"></td>
						<td align=center><INPUT type="text" size=12 class='ratakanan' name='ffacpo1' value='0' onkeypress="return harusangka(event);"></td>
						<td align=center><INPUT type="text" size=12 class='ratakanan' name='aircpo1' value='0' onkeypress="return harusangka(event);"></td>
						<td align=center><INPUT type="text" class='ratakanan' size=12 name='kotorancpo1' value='0' onkeypress="return harusangka(event);"></td>
					</tr>
					<tr>
						<td align=center><INPUT type="text" size=12 name='TankID2' value='2' readonly></td>
						<td align=center><INPUT type="text" size=12 class='ratakanan' name='isicpo2' id='isicpo2' value='0' onkeyup="getTotal()" onkeypress="return harusangka(event);"></td>
						<td align=center><INPUT type="text" size=12 class='ratakanan' name='ffacpo2' value='0' onkeypress="return harusangka(event);"></td>
						<td align=center><INPUT type="text" size=12 class='ratakanan' name='aircpo2' value='0' onkeypress="return harusangka(event);"></td>
						<td align=center><INPUT type="text" class='ratakanan' size=12 name='kotorancpo2' value='0' onkeypress="return harusangka(event);"></td>
					</tr>
					<tr>
						<td align=center><INPUT type="text" size=12 name='TankID3' value='3' readonly></td>
						<td align=center><INPUT type="text" size=12 class='ratakanan' name='isicpo3' id='isicpo3' value='0' onkeyup="getTotal()" onkeypress="return harusangka(event);"></td>
						<td align=center><INPUT type="text" size=12 class='ratakanan' name='ffacpo3' value='0' onkeypress="return harusangka(event);"></td>
						<td align=center><INPUT type="text" size=12 class='ratakanan' name='aircpo3' value='0' onkeypress="return harusangka(event);"></td>
						<td align=center>
							<INPUT type="text" class='ratakanan' size=12 name='kotorancpo3' value='0' onkeypress="return harusangka(event);">
						</td>
					</tr>	
					<tr>
						<td>
							Total
						</td>
						<td align=left colspan="4">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT type="text" class='ratakanan' size=12 name='totalcpo' readonly id='totalcpo' value='0'>
						</td>
					</tr>
				</table>
			</td>
		</tr>

		<tr>
			<td align="left" class="lebarkolomsatu">Adjustment CPO</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" name="adjcpo" id="adjcpo" size="15" value="0" onkeypress=""> KG
			</td>
			
		</tr>
	
		<tr>
			<td align="left">C.O.T</td>
			<td align="left">
				<input type="text" class="ratakanan" name="cotcpo" id="cotcpo" size="15" value="0" onkeypress="return harusangka(event);"> KG
			</td>
		</tr>
		<tr>
			<td align="left">C.S.T</td>
			<td align="left">
				<input type="text" class="ratakanan" name="cstcpo" id="cstcpo" size="15" value="0" onkeypress="return harusangka(event);"> KG
			</td>
		</tr>
		<tr>
			<td align="left">O.T</td>
			<td align="left">
				<input type="text" class="ratakanan" name="otcpo" id="otcpo" size="15" value="0" onkeypress="return harusangka(event);"> KG
			</td>
		</tr>
		<tr>
			<td align="left">FFA (%)</td>
			<td align="left">
				<input type="text" class="ratakanan" name="ffacpo" id="ffacpo" size="15" value="0" onkeypress="return harusangka(event);"> KG
			</td>
		</tr>
		<tr>
			<td align="left">AIR (%)</td>
			<td align="left">
				<input type="text" class="ratakanan" name="aircpo" id="aircpo" size="15" value="0" onkeypress="return harusangka(event);"> KG
			</td>
		</tr>
		<tr>
			<td align="left">KOTORAN (%)</td>
			<td align="left">
				<input type="text" class="ratakanan" name="kotorancpo" id="kotorancpo" size="15" value="0" onkeypress="return harusangka(event);"> KG
			</td>
		</tr>
		<tr>
			<td align="left">DOBI (%)</td>
			<td align="left">
				<input type="text" class="ratakanan" name="dobicpo" id="dobicpo" size="15" value="0" onkeypress="return harusangka(event);"> KG
			</td>
		</tr>
		<tr bgcolor="#00a300" style="color:#FFFFFF;" style="color:#FFFFFF;"">
			<td valign="top" align="center" colspan="4" >INTI KELAPA SAWIT</td>
		</tr>
		<tr>
			<td align="left" class="lebarkolomsatu">Kirim</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" name="kirimkernel" id="kirimkernel" readOnly size="15" value="<%=NETTOKERNEL%>"> KG
			</td>
			<td colspan="2" rowspan="2">
				<table class="data warnagaris" align=left border="1"  width="100%" id='table2'>
					<tr >
						<td align=center class="tebal">Bulking Silo</td>
						<td align=center class="tebal">Isi (Kg)</td>
						<td align=center class="tebal">FFA (%)</td>
						<td align=center class="tebal">Air (%)</td>
						<td align=center class="tebal">Kotoran (%)</td>
					</tr>
					<tr>
						<td align=center><INPUT type="text" size=12 name='bulkingsilo' value='1' readonly onkeypress="return harusangka(event);"></td>
						<td align=center><INPUT type="text" size=12 class='ratakanan' name='isibulking' value='0' onkeypress="return harusangka(event);"></td>
						<td align=center><INPUT type="text" size=12 class='ratakanan' name='ffabulking' value='0' onkeypress="return harusangka(event);"></td>
						<td align=center><INPUT type="text" size=12 class='ratakanan' name='airbulking' value='0' onkeypress="return harusangka(event);"></td>
						<td align=center><INPUT type="text" class='ratakanan' size=12 name='kotoranbulking' value='0' onkeypress="return harusangka(event);"></td>
					</tr>
					
				</table>
			</td>
		</tr>
		<tr>
			<td align="left" class="lebarkolomsatu">Kernel Silo</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" name="silokernel" id="silokernel" size="15" value="0" onkeypress="return harusangka(event);"> KG
			</td>
		</tr>

		<tr bgcolor="#00a300" style="color:#FFFFFF;" style="color:#FFFFFF;"">
			<td valign="top" align="center" colspan="4" >CANGKANG</td>
		</tr>
		<tr>
			<td align="left" class="lebarkolomsatu">Kirim</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" name="kirimckg" id="kirimckg" readOnly size="15" value="<%=nettockg%>"> KG
			</td>
		</tr>
		<tr>
			<td align="left" class="lebarkolomsatu">Produksi</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" name="produksickg" id="produksickg" size="15" value="0" onkeypress="return harusangka(event);"> KG
			</td>
		</tr>
		<tr bgcolor="#00a300" style="color:#FFFFFF;" style="color:#FFFFFF;"">
			<td valign="top" align="center" colspan="4" >TENAGA KERJA</td>
		</tr>
		<tr>
			<td colspan="4" >
				<table class="data warnagaris" align=center border="1"  width="100%" id='table1'>
					<tr >
						<td align=center class="tebal">SHIFT</td>
						<td align=center class="tebal">HK</td>
						<td align=center class="tebal">CUTI</td>
						<td align=center class="tebal">SAKIT</td>
						<td align=center class="tebal">IZIN/P1</td>
						<td align=center class="tebal">MANGKIR</td>
					</tr>
					<tr>
						<td align=center><INPUT type="text" class="ratatengah" size=12 name='shift1' value='1' readonly></td>
						<td align=center><INPUT type="text" size=12 class='ratakanan' name='hk1' value='<%=hk1%>' readonly></td>
						<td align=center><INPUT type="text" size=12 class='ratakanan' name='cuti1' value='<%=cuti1%>' readonly></td>
						<td align=center><INPUT type="text" size=12 class='ratakanan' name='sakit1' value='<%=sakit1%>' readonly></td>
						<td align=center><INPUT type="text" class='ratakanan' size=12 name='izin1' value='<%=izin1%>' readonly></td>
						<td align=center><INPUT type="text" class='ratakanan' size=12 name='mangkir1' value='<%=mangkir1%>' readonly></td>
					</tr>	
					<tr>
						<td align=center><INPUT type="text" size=12 name='shift2' class="ratatengah" value='2' readonly></td>
						<td align=center><INPUT type="text" size=12 class='ratakanan' name='hk2' value='<%=hk2%>' readonly></td>
						<td align=center><INPUT type="text" size=12 class='ratakanan' name='cuti2' value='<%=cuti2%>' readonly></td>
						<td align=center><INPUT type="text" size=12 class='ratakanan' name='sakit2' value='<%=sakit2%>' readonly></td>
						<td align=center><INPUT type="text" class='ratakanan' size=12 name='izin2' value='<%=izin2%>' readonly></td>
						<td align=center><INPUT type="text" class='ratakanan' size=12 name='mangkir2' value='<%=mangkir2%>' readonly></td>
					</tr>	
					
				</table>
			</td>
		</tr>
		<tr bgcolor="#00a300" style="color:#FFFFFF;" style="color:#FFFFFF;"">
			<td valign="top" align="center" colspan="4" >PEMAKAIAN AIR dan CaCo3</td>
		</tr>
		
		<tr>
			<td align="left" class="lebarkolomsatu">AIR</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" name="air" id="air" size="15"  value="0" onkeypress="return harusangka(event);"> M3
			</td>
			
		</tr>
		<tr>
			<td align="left" class="lebarkolomsatu">CaCo3</td>
			<td align="left" class="lebarkolomdua">
				<input type="text" class="ratakanan" name="caco3" id="caco3" size="15" value="0" onkeypress="return harusangka(event);"> KG
			</td>
		</tr>
		<tr bgcolor="#00a300" style="color:#FFFFFF;" style="color:#FFFFFF;"">
			<td valign="top" align="center" colspan="4" >KETERANGAN</td>
		</tr>
		<tr>
			<td valign="top" align="center" colspan="4" >
				<textarea name="ket" cols="150" rows="3" placeholder="isi keterangan"></textarea>
			</td>
		</tr>
		<tr>
			<td align="center" colspan="6">
		        <input Type="Submit" name="Add" value="Simpan">
			</td>
		</tr>
	
	</table>
</form>
</div></div>
  <!--#include file="bottom.jsp"-->

	</body>
	<!-- xx-->
</html>
