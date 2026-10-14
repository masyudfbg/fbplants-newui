<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim businessunit,fcbaname,CompanyName,q,sqltotdata,gangcode
businessunit= session("fcba")
gangcode	= trim(request("gangcode"))

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
 
end if



%>

<html>
	<head>
		<title>FBPlants - BKM TRAKSI</title>
		<link href="/js/tabbedpanels/SpryTabbedPanels.css" rel="stylesheet" type="text/css"/>
		<script src="/js/tabbedpanels/SpryTabbedPanels.js" type="text/javascript"></script>
		<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>
		<link rel="stylesheet" href="css/jquery-ui.css" />
		<link rel="StyleSheet" href="style-content.css" type="text/css" />

		<link href="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.css" rel="stylesheet" type="text/css"/>
		<script src="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.js" type="text/javascript"></script>
		
		<link href="/js/spry/widgets/selectvalidation/SpryValidationSelect.css" rel="stylesheet" type="text/css"/>
		<script src="/js/spry/widgets/selectvalidation/SpryValidationSelect.js" type="text/javascript"></script>

	<% if Request("oldui") <> "1" then %>
<!-- Tampilan baru (css/fbp-shell.css + css/fbp-forms.css). ?oldui=1 = tampilan lama -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="StyleSheet" href="/css/login/font-awesome.min.css" type="text/css" />
<link rel="StyleSheet" href="/css/fbp-shell.css?v=20261009d" type="text/css" />
<link rel="StyleSheet" href="/css/fbp-forms.css?v=20261009e" type="text/css" />
<script type="text/javascript" src="/js/fbp-shell.js?v=20261014b"></script>
<% end if %>
</head>

<style>
	.posisikirijudul{
		position: relative;
		left: 0px;
	}
	.posisikiri{
		position: relative;
		left: 15px;
	}
	.lebarsection{
		width: 160px;
	}
	.lebargangcode{
		width: 200px;
	}
	.lebarhk{
		width: 45px;
	}
	.lebarabsen{
		width: 45px;
	}
	.lebarlokasi{
		width: 50px;
	}
	.lebartc{
		width: 60px;
	}
	.lebarovtime{
		width: 70px;
	}
	.lebarrate{
		width: 60px;
	}
	.lebaractivity{
		width: 110px;
	}
	.lebarallow{
		width: 90px;
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
	.lebaritembarang{
		width: 400px;
	}
	.lebaractbarang{
		width: 300px;
	}
	.lebarvehicle{
		width: 200px;
	}
	.lebartime{
		width: 90px;
	}

	.lebarspkno{
		width: 250px;
	}
	
</style>

<script>

</script>

<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy page-bkm"<% end if %>>
<!--#include file="top.jsp"-->

<br>
<br>
<br>
<br>
<br>


<div id="head_judul" style="float:left;font-size:10px; text-align:left;">
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : BKM TRAKSI</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry" method="post" action="bkm_traksi_proses.asp" onsubmit="return false;">
	
	<input type="hidden" name="jumldata" id="jumldata" value="0" />
	<input type="hidden" name="jumldatavh" id="jumldatavh" value="1" />
	<input type="hidden" name="proses" id="proses" value="1" />
	<input type="hidden" name="fcbadata" id="fcbadata" value=<%=businessunit%> />
	<table class="data posisikiri" border="0" width="90%">			
		<tr>
			<td align="left" >&nbsp;&nbsp;&nbsp;Tanggal</td>
			<td align="left">
				<input type="text" class="datepicker" maxlength="15" name="tglbkm" readonly id="tglbkm" size="15" value="<%=textDate(date)%>" title="Tanggal BKM" onchange="get_Allowentry();">
			</td>
			<td align="left" >Section Name</td>
			<td align="left">
				<select name="sectionname" id="sectionname" class="lebarsection" onchange="showGang();">
				</select>
			</td>
			<td align="left" >&nbsp;&nbsp;&nbsp;Mandor</td>
			<td align="left">
				<select name="supervisi" id="supervisi" class="lebarsection" >
				</select>
			</td>
			
		
		</tr>
		<tr>
			<td align="left" >&nbsp;&nbsp;&nbsp;Gang</td>
			<td align="left">
				<select name="gangcode" id="gangcode" class="lebargangcode" >
				</select>
			</td>
			<td align="left" >Kendaraan</td>
			<td align="left">
				<select name="kendaraan" id="kendaraan" class="lebarsection" >
				</select>
			</td>
			
		</tr>
	</table>
	
	<br>
		
	<table align=center border="0" class='data posisikiri' cellpadding="0" cellspacing="0" width='1100' id='tableall'>
		<tr>
			<td colspan=>
				<div class="TabbedPanels" id="tp1" align='center'>
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">BKM UPAH</li>
						<li class="TabbedPanelsTab">BKM VEHICLE/WORKSHOP</li>
						
					</ul>

					<div class="TabbedPanelsContentGroup">

						<div class="TabbedPanelsContent">
							<table align=center border="0" width="1200" class='data' id='table2'>
								<tr id="trheader2">
									<th>Employee</th>
									<th>Status</th>
									<th>TimeIn</th>
									<th>TimeOut</th>
									<th>Lokasi</th>
									<th>Absen</th>
									<th>Activity</th>
									<th>TC</th>
									<th>HK</th>
									<th>Unit</th>
									<th>Output</th>
									<th>Rate</th>
									<th>Allowed Code</th>
									<th>Rate Allowed</th>
									<th>OVTH</th>
									<th>TipeOVTime</th>
									<th>Remark</th>
									
									
									<th colspan="2">&nbsp;</th>
								</tr>
								<tr id='tru'>
									<tbody id="bodytable">
									</tbody>
								</tr>
								
							</table>
						</div>

						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="datadetail">
								<table align=center border="0"  width="1050" class='data' id='table3'>
									<tr id="trheader3">
										<th>Opr/Driver</th>
										<th>Helper</th>
										<th>Lokasi</th>
										<th>Time</th>
										<th>Activity</th>
										<th>TC</th>
										<th>Category Transport</th>
										<th>Trip</th>
										<th>Hours</th>
										<th>KM/HM</th>
										<th>Output</th>
										<th>Remark</th>
										<th colspan="2">&nbsp;</th>
									</tr>


										<tr id='trv'>
											<tbody id="bodytablev">
											</tbody>
											
										</tr>
									
								</table>
							</div>
							<!-- ============================================================================================================================= -->
						</div>

					</div>
				</div>
			</td>
		</tr>
		<tr>
			<td colspan='' align=center><input Type="Submit"  id="btnsimpan" name="simpan" value="Simpan" onclick="prosesSubmit();"></td>
		</tr>
	</table>

</form>
</div>
</body>
<!--#include file="bottom.jsp"-->
</html>
<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
<script src="/js/select2/select2.min.js" type="text/javascript"></script>

<script type="text/javascript">
	//window.onload = getcountEmployee();
	new Spry.Widget.TabbedPanels("tp1", {
		defaultTab: 0
	});

	function hanyaAngka(evt) {
		evt.value = evt.value.replace(/[^0-9,.-]/g,"");
	}

$(function() {	
	$("#tglbkm").datepicker();
	showSection();
	showVH();
	showSupervisi2();
	//addrowVH(1);
	
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


//get section
function showSection() {
	
	$('#sectionname').empty();
	$('#sectionname').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getsection.asp?section=' + 1,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//get GANGCODE 
function showGang() {
	var sectionname = $("#sectionname").val();
	$('#gangcode').empty();
	$('#gangcode').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_gangcode.asp?sectionname=' + sectionname + '&gancode=' + 1,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//get supervisi 2
function showSupervisi2() {
	
	$('#supervisi').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getsupervisi.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//get absen
function showAbsen(selid) {
	
	$('#absen').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getabsen.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//get lokasi upah
function showFCBA(selid) {
	var fcba = 'fcba';
	$('#lokasi').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getabsen.asp?viewdata=' + fcba,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//show fcba vh
function showFCBAVH(selid) {
	var fcba = 'fcba';
	$('#lokasivh').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getabsen.asp?viewdata=' + fcba,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

//show kendaraan
function showVH() {
	$('#kendaraan').empty();
	var jmlitem = 2;
	$('#kendaraan').select2({
		ajax: {
			delay: 1000,
			url : "bkm_show_item.asp?jmlitem=" + jmlitem,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}


//get activity
function showActivity(selid) {
	
	$('#activity').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp?panen=' + 2,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}


//get allowedcode
function showAllowed(selid) {
	
	$('#allowedcode').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getallowed.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}

function get_Allowentry()
{
	var gangcode = $("#gangcode").val();
	var tglbkm = $('#tglbkm').val();
	if (tglbkm != "") {
		$.post("bkm_allow_entry.asp?tglbkm=" + tglbkm, function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var closingStatus = json.results[0].msg;
			if(closingStatus=="FALSE"){
				getEmployee();
				addrowVH(1);
				$('#btnsimpan').show();
				//window.open('bkm_operasional_edit.Asp?bkmno=' + bkmno);
			}else{
				alert("PERIODE SUDAH DI CLOSE !!!");
				$('#btnsimpan').hide();
				//break;
			}
			
			
						
		});
		
	} 
}

function getEmployee()
{
	var gangcode = $("#gangcode").val();
	var vehiclecode = $("#kendaraan").val();
	var tglbkm = $('#tglbkm').val();
	if (gangcode != "") {
		$.post("bkm_employee.asp?gangcode=" + gangcode + "&vehiclecode=" + vehiclecode + "&tglbkm=" + tglbkm + "&employee=" + 200, function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var jumlview  = json.results.length;
			if(jumlview>0){
				var bkmno = json.results[0].bkmno;
				window.close();
				window.open('bkm_traksi_edit.Asp?bkmno=' + bkmno);
			}else{
				addrowUpah('1');
			}
			
			
						
		});
		
	} 
}

function prosesSubmit() {
	
	if($('#fcbadata').val() == "ARM"){
		document.frmEntry.submit();
	}else{
		var errors = [];

		// === VALIDASI HEADER ===
		if (!$.trim($('#tglbkm').val()))    errors.push("• Tanggal BKM harus dipilih.");
		if (!$('#sectionname').val())        errors.push("• Section Name harus dipilih.");
		if (!$('#kendaraan').val())          errors.push("• Kendaraan harus dipilih.");

		if (errors.length > 0) {
			alert("Inputan anda belum lengkap:\n\n" + errors.join("\n"));
			return;
		}

		// === TENTUKAN TAB YANG SEDANG AKTIF ===
		// index 0 = BKM UPAH, index 1 = BKM VEHICLE/WORKSHOP
		var activeTab = $('#tp1 .TabbedPanelsTab').index($('#tp1 .TabbedPanelsTabSelected'));

		if (activeTab === 0) {
			// === VALIDASI TAB BKM UPAH ===
			var jumldata = parseInt($('#jumldata').val()) || 0;
			if (jumldata === 0) {
				alert("Inputan anda belum lengkap:\n\n• Tab BKM UPAH belum diisi.");
				return;
			}

			var upahErrors = [];
			var skipValidasi = ["KJ", "WS", "WH", "DL", "KH"];

			for (var i = 1; i <= jumldata; i++) {
				if ($('#tru' + i).length === 0) continue;

				if (!$('#fccode'   + i).val()) upahErrors.push("Baris " + i + " (Upah): Employee harus diisi.");
				if (!$('#lokasi'   + i).val()) upahErrors.push("Baris " + i + " (Upah): Lokasi harus diisi.");
				if (!$('#absen'    + i).val()) upahErrors.push("Baris " + i + " (Upah): Absen harus diisi.");
				if (!$('#activity' + i).val()) upahErrors.push("Baris " + i + " (Upah): Activity harus diisi.");
				if (!$('#tcupah'   + i).val()) upahErrors.push("Baris " + i + " (Upah): TC harus diisi.");

				var absen        = $('#absen'        + i).val();
				var presentgrade = $('#presentgrade' + i).val();

				if (!skipValidasi.includes(absen)) {
					var hk = parseFloat($.trim($('#hk' + i).val())
						.replace(/\./g,"")
						.replace(",","."));

					if (isNaN(hk) || hk <= 0)
						upahErrors.push("Baris " + i + " (Upah): HK harus lebih dari 0.");
				}

				// Validasi tambahan: jika Status KHT, HK wajib diisi > 0
				if (presentgrade === "KHT") {
					var hkKHT = parseFloat($.trim($('#hk' + i).val())
						.replace(/\./g,"")
						.replace(",","."));

					if (isNaN(hkKHT) || hkKHT <= 0)
						upahErrors.push("Baris " + i + " (Upah): Status KHT, HK harus lebih dari 0.");
				}

				var ovtime = parseFloat($.trim($('#ovtimehours' + i).val()).replace(/\./g,"").replace(",","."));
				if (!isNaN(ovtime) && ovtime > 0 && !$('#typeovrtime' + i).val()) {
					upahErrors.push("Baris " + i + " (Upah): Tipe Overtime harus dipilih.");
				}
			}

			if (upahErrors.length > 0) {
				alert("Inputan anda belum lengkap:\n\n" + upahErrors.join("\n"));
				return;
			}
		} else {
			// === VALIDASI TAB BKM VEHICLE ===
			var jumldatavh = parseInt($('#jumldatavh').val()) || 0;
			if (jumldatavh === 0) {
				alert("Inputan anda belum lengkap:\n\n• Tab BKM VEHICLE/WORKSHOP belum diisi.");
				return;
			}

			var vhErrors = [];
			for (var j = 1; j <= jumldatavh; j++) {
				if ($('#trv' + j).length === 0) continue;

				if (!$('#employeecode' + j).val()) vhErrors.push("Baris " + j + " (Vehicle): Operator/Driver harus diisi.");
				if (!$('#lokasivh'    + j).val())  vhErrors.push("Baris " + j + " (Vehicle): Lokasi harus diisi.");
				if (!$('#activityivh' + j).val())  vhErrors.push("Baris " + j + " (Vehicle): Activity harus diisi.");
				if (!$('#tcvh'        + j).val())  vhErrors.push("Baris " + j + " (Vehicle): TC harus diisi.");

				var timestart = $.trim($('#timestart' + j).val());
				if (!timestart) {
					vhErrors.push("Baris " + j + " (Vehicle): Time harus diisi.");
				} else if (!/^\d{2}:\d{2}-\d{2}:\d{2}$/.test(timestart)) {
					vhErrors.push("Baris " + j + " (Vehicle): Format Time harus hh:mm-hh:mm.");
				}

				var jamkerja = parseFloat($.trim($('#jamkerja' + j).val()).replace(/\./g,"").replace(",","."));
				if (isNaN(jamkerja) || jamkerja <= 0) vhErrors.push("Baris " + j + " (Vehicle): Hours harus lebih dari 0.");
			}

			if (vhErrors.length > 0) {
				alert("Inputan anda belum lengkap:\n\n" + vhErrors.join("\n"));
				return;
			}
		}

		// === SEMUA LOLOS, BARU SUBMIT ===
		document.frmEntry.submit();
    }
}


//getEmployee
function getEmployeeview() {
	
	var gangcode = $("#gangcode").val();
	var tglbkm = $('#tglbkm').val();
	
	if (gangcode != "") {
		$.post("bkm_employee.asp?gangcode=" + gangcode + "&tglbkm=" + tglbkm + "&employee=" + 1, function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var jumlview  = json.results.length;
			var no = 1;
			for (var i = 0; i < jumlview; i++) {
				var fccode = json.results[i].fccode;
				var fcname = json.results[i].fcname;
				var fcba = json.results[i].fcba;
				var dttimein = json.results[i].timeinx;
				var dttimeout = json.results[i].timeoutx;
				var datatimein = json.results[i].timein;
				var datatimeout = json.results[i].timeout;
				var datapresentgrade = json.results[i].presentgrade;
						
				
				var docketlist = '';
					docketlist += '<tr id="tru'+no+'">';
					docketlist += '<td align="left"  class="ratakiri lebaractivity "><select class="ratakiri lebaractivity" name="fccode' + no + '" id="fccode' + no + '" onchange="getFinger(' + no + ');"><option value=' + fccode + '>' + fcname + '</option></select></td>';
					docketlist += '<td align="center">' + datapresentgrade + '<input type="hidden" class="ratatengah lebarhk" name="presentgrade' + no + '" id="presentgrade' + no + '" value="' + datapresentgrade + '"/></td>';
					docketlist += '<td align="center">' + datatimein + '<input type="hidden" class="ratatengah lebarhk" name="timeinx' + no + '" id="timeinx' + no + '" value="' + dttimein + '"/></td>';
					docketlist += '<td align="center">' + datatimeout + '<input type="hidden" class="ratatengah lebarhk" name="timeoutx' + no + '" id="timeoutx' + no + '" value="' + dttimeout + '"/></td>';
					docketlist += '<td align="left" class="ratakiri lebarabsen"><select class="ratakiri lebarlokasi" name="lokasi' + no + '" id="lokasi' + no + '"><option value=' + fcba + '>' + fcba + '</option></select></td>';
					docketlist += '<td align="left" class="ratakiri lebarabsen"><select class="ratakiri lebarabsen" name="absen' + no + '" id="absen' + no + '" onchange="getDisable(' + no + ')"/></select></td>';
					docketlist += '<td align="center" class="ratakiri lebaractivity"><select class="ratakiri lebaractivity" name="activity' + no + '" id="activity' + no + '" onchange="getTC(' + no + ');"/></select></td>';
					docketlist += '<td align="center"><select class="ratatengah lebartc" name="tcupah' + no + '" id="tcupah' + no + '"/></select></td>';		
					docketlist += '<td align="center"><input class="ratatengah lebarhk" type="text" onkeyup="hanyaAngka(this)" name="hk' + no + '" id="hk' + no + '" value="0"/></td>';
					docketlist += '<td align="center"><input class="ratatengah lebarabsen" type="text" onkeyup="hanyaAngka(this)" name="unit' + no + '" id="unit' + no + '" value="0"/></td>';
					docketlist += '<td align="center"><input class="ratatengah lebarabsen" type="text" onkeyup="hanyaAngka(this)" name="output' + no + '" id="output' + no + '" value="0"/></td>';
					docketlist += '<td align="center"><input class="ratatengah lebarrate" type="text" onkeyup="hanyaAngka(this)" name="ratekb' + no + '" id="ratekb' + no + '" value="0"/></td>';		
					docketlist += '<td align="center" class="ratakiri lebaractivity"><select class="ratakiri lebarallow" name="allowedcode' + no + '" id="allowedcode' + no + '"/></select></td>';
					docketlist += '<td align="center"><input class="ratatengah lebarrate" type="text" onkeyup="hanyaAngka(this)" name="ratepremi' + no + '" id="ratepremi' + no + '" value="0"/></td>';		
					docketlist += '<td align="center"><input class="ratatengah lebarabsen" type="text" onkeyup="hanyaAngka(this);getOTHConvertion(' + no + ');" name="ovtimehours' + no + '" id="ovtimehours' + no + '" value="0"/></td>';		
					docketlist += '<td align="center"><select class="ratatengah lebarovtime" onchange="getOTHConvertion(' + no + ');" name="typeovrtime' + no + '" id="typeovrtime' + no + '"/></select><input class="ratatengah" type="hidden" onkeyup="hanyaAngka(this)" name="othconvertion' + no + '" id="othconvertion' + no + '" value="0"/></td>';		
					docketlist += '<td align="center"><textarea rows="3" cols="10" name="remark' + no + '" id="remark' + no + '"/></textarea></td>';		
					docketlist += '<td align="center"><img src="css/login/images/add.jpg" onclick="addrowUpah(' + no + ',' + no + ')"/></td>';
					docketlist += '<td align="center"><img src="css/login/images/del.png" onclick="deleterowUpah(' + no + ')"/></td>';
					docketlist += '</tr>';
					$('#bodytable').append(docketlist);
				
				
				showEmployee("fccode" + no);
				showFCBA("lokasi" + no);
				showAbsen("absen" + no);
				showActivity("activity" + no);
				showAllowed("allowedcode" + no);
				showTypeOv("typeovrtime" + no);

				new Spry.Widget.ValidationSelect("fccode" + no);
				new Spry.Widget.ValidationSelect("presentgrade" + no);
				new Spry.Widget.ValidationSelect("lokasi" + no);
				new Spry.Widget.ValidationSelect("absen" + no);
				new Spry.Widget.ValidationSelect("activity" + no);
				new Spry.Widget.ValidationSelect("tcupah" + no);
				new Spry.Widget.ValidationTextField("hk" + no);
				
				no++;
				$("#jumldata").val(jumlview);
			}
			
			
			
			
		});
		
	} else {
		$('#bodytable').empty();
	}
	
}


function getTC(i)
{
	var activity = $('#activity' + i).val();
	var lokasi   = $('#lokasi' + i).val();

	$('#tcupah' + i).empty();
	$('#tcupah' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getblok.asp?activity=' + activity +'&lokasi=' + lokasi,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});

}


function getTCVH(i)
{
	var activityivh = $('#activityivh' + i).val();
	var lokasi   = $('#lokasivh' + i).val();
		var tglbkm = $('#tglbkm').val();
	$('#tcvh' + i).empty();
	$('#tcvh' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getblok.asp?activity=' + activityivh + '&lokasi=' + lokasi + '&tglbkm=' + tglbkm,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});

}


//show tipe overtime
function showTypeOv(selid)
{
	
	//$('#typeovrtime' + i).empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_viewdata.asp?idview=' + 7,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});

}

//view data oth convertion
function getOTHConvertion(i)
{
	var ovtimehours = $("#ovtimehours" + i).val();
	var typeovrtime = $('#typeovrtime' + i).val();
	if (ovtimehours != "") {
		$.post("bkm_employee.asp?ovtimehours=" + ovtimehours + "&typeovrtime=" +  typeovrtime +  "&satuan=" + 7, function(data) {
			$('#othconvertion' + i).empty();
			var json  = JSON.parse(data);
			var othconvertion = json.results[0].convertion
			$('#othconvertion' + i).val(othconvertion);
						
		});
		
	}

	
}

//get employee
function showEmployee(selid) {
	var gangcode = $("#gangcode").val();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : "bkm_employee_edit.asp?gangcode=" + gangcode,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
	//alert(gangcode);
}



var StartIteration=1;
function addrowUpah(trid) {
	

	// var lastrow	= parseInt(nourut)+1;
	 var lastrow = this.StartIteration++;
	
	
	if (lastrow < 2) {
		var row = $("#trheader2").after("<tr id='tru" + lastrow + "'></tr>");
		var kurangbutton = "&nbsp;";
	} else {
		var row = $('#tru' + trid).after("<tr id='tru" + lastrow + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowUpah(" + lastrow + ");' />";
	}

	var fccode = "<select name='fccode" + lastrow + "' id='fccode" + lastrow + "' class='ratakiri lebaractivity' onchange='getFinger(" + lastrow + ");'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri'>" + fccode + "</td>");

	var presentgrade = "<span name='presentgrade" + lastrow + "' id='presentgrade" + lastrow + "'></span><input type='hidden' name='presentgrade" + lastrow + "' id='presentgrade_val" + lastrow + "'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + presentgrade + "</td>");

	var timein = "<span name='timein" + lastrow + "' id='timein" + lastrow + "' class='ratatengah lebarrate'></span><input type='hidden' name='timeinx" + lastrow + "' id='timeinx" + lastrow + "'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + timein + "</td>");

	var timeout = "<span name='timeout" + lastrow + "' id='timeout" + lastrow + "' class='ratatengah lebarrate'></span><input type='hidden' name='timeoutx" + lastrow + "' id='timeoutx" + lastrow + "'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + timeout + "</td>");

	var lokasi = "<select name='lokasi" + lastrow + "' id='lokasi" + lastrow + "' class='ratakiri lebarlokasi'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri lebarlokasi'>" + lokasi + "</td>");

	var absen = "<select name='absen" + lastrow + "' id='absen" + lastrow + "' class='ratakiri lebarabsen' onchange='getDisable(" + lastrow + ");'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri lebarabsen'>" + absen + "</td>");

	var activity = "<select name='activity" + lastrow + "' id='activity" + lastrow + "' class='ratakiri lebaractivity' onchange='getTC(" + lastrow + ");getDisable(" + lastrow + ");'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri lebarabsen'>" + activity + "</td>");

	var blok = "<select type='text' name='tcupah" + lastrow + "' id='tcupah" + lastrow + "' value='0' class='ratatengah lebartc'></select>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + blok + "</td>");

	var hk = "<input type='text' name='hk" + lastrow + "' id='hk" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarhk'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + hk + "</td>");

	var unit = "<input type='text' name='unit" + lastrow + "' id='unit" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + unit + "</td>");

	var output = "<input type='text' name='output" + lastrow + "' id='output" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + output + "</td>");

	var ratekb = "<input type='text' name='ratekb" + lastrow + "' id='ratekb" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarrate'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + ratekb + "</td>");

	var allowedcode = "<select name='allowedcode" + lastrow + "' id='allowedcode" + lastrow + "' class='ratakiri lebarallow'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri'>" + allowedcode + "</td>");

	var ratepremi = "<input type='text' name='ratepremi" + lastrow + "' id='ratepremi" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarrate'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + ratepremi + "</td>");

	var ovtimehours = "<input type='text' name='ovtimehours" + lastrow + "' id='ovtimehours" + lastrow + "' onkeyup='hanyaAngka(this);getOTHConvertion(" + lastrow + ")' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + ovtimehours + "</td>");

	var typeovrtime = "<select type='text' name='typeovrtime" + lastrow + "' id='typeovrtime" + lastrow + "' onchange='getOTHConvertion(" + lastrow + ");' value='0' class='ratatengah lebarovtime'></select><input type='hidden' name='othconvertion" + lastrow + "' id='othconvertion" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + typeovrtime + "</td>");

	var remark = "<textarea type='text' name='remark" + lastrow + "' id='remark" + lastrow + "' rows='3' cols='10' ></textarea>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + remark + "</td>");


	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowUpah(" + lastrow + ");' />";
	$('#tru' + lastrow).append("<td>" + tambahbutton + "</td>");

	$('#tru' + lastrow).append("<td>" + kurangbutton + "</td>");

	$('#jumldata').val(lastrow);
	

	showEmployee("fccode" + lastrow);
    showFCBA("lokasi" + lastrow);
	showAbsen("absen" + lastrow);
	showActivity("activity" + lastrow);
	showAllowed("allowedcode" + lastrow);
	showTypeOv("typeovrtime" + lastrow);

	new Spry.Widget.ValidationSelect("fccode" + lastrow);
	new Spry.Widget.ValidationSelect("lokasi" + lastrow);
	new Spry.Widget.ValidationSelect("absen" + lastrow);
	new Spry.Widget.ValidationSelect("activity" + lastrow);
	new Spry.Widget.ValidationSelect("tcupah" + lastrow);
	new Spry.Widget.ValidationTextField("hk" + lastrow);
	

}

function deleterowUpah(i) {
	Spry.Widget.Utils.destroyWidgets('fccode' + i);
	Spry.Widget.Utils.destroyWidgets('presentgrade' + i);
	Spry.Widget.Utils.destroyWidgets('timein' + i);
	Spry.Widget.Utils.destroyWidgets('timeout' + i);
	Spry.Widget.Utils.destroyWidgets('timeinx' + i);
	Spry.Widget.Utils.destroyWidgets('timeoutx' + i);
	Spry.Widget.Utils.destroyWidgets('lokasi' + i);
	Spry.Widget.Utils.destroyWidgets('absen' + i);
	Spry.Widget.Utils.destroyWidgets('activity' + i);
	Spry.Widget.Utils.destroyWidgets('tcupah' + i);
	Spry.Widget.Utils.destroyWidgets('hk' + i);
	Spry.Widget.Utils.destroyWidgets('unit' + i);
	Spry.Widget.Utils.destroyWidgets('output' + i);
	Spry.Widget.Utils.destroyWidgets('ratekb' + i);
	Spry.Widget.Utils.destroyWidgets('allowedcode' + i);
	Spry.Widget.Utils.destroyWidgets('ratepremi' + i);
	Spry.Widget.Utils.destroyWidgets('ovtimehours' + i);
	Spry.Widget.Utils.destroyWidgets('typeovrtime' + i);
	Spry.Widget.Utils.destroyWidgets('othconvertion' + i);
	Spry.Widget.Utils.destroyWidgets('remark' + i);
	
	$('#tru' + i).remove();	
	
}

//show time from finger employee
function getFinger(i)
{
	var tglbkm = $('#tglbkm').val();
	var fccode = $('#fccode' + i).val();
	//alert(tglbkm);
	if (fccode != "") {
		$.post("bkm_employee_edit.asp?emp=" + 1 + "&tglbkm=" + tglbkm + "&fccode=" + fccode, function(data) {
			var json  = JSON.parse(data);
			//var jumlview  = json.results.length;
			var timein = json.results[0].timein;
			var timeinx = json.results[0].timeinx;
			var timeout = json.results[0].timeout;
			var timeoutx = json.results[0].timeoutx;
			var datapresentgrade = json.results[0].presentgrade;
			$('#timein' + i).html(timein);
			$('#timeout' + i).html(timeout);
			$('#presentgrade' + i).html(datapresentgrade);
			$('#timeinx' + i).val(timeinx);
			$('#timeoutx' + i).val(timeoutx);			
			$('#presentgrade' + i).val(datapresentgrade);			
		});
		
	}

}


//function disabled input()
function getDisable(i)
{
	var absen = $('#absen' + i).val();
	if (absen=="KB"){
		//disbale
		$( '#hk' + i ).prop( "disabled", true );
		$( '#ratepremi' + i ).prop( "disabled", true );
		$( '#allowedcode' + i ).prop( "disabled", true );
		
		//enable
	
	}else{
		$( '#hk' + i ).prop( "disabled", false );
		$( '#ratepremi' + i ).prop( "disabled", false );
		$( '#unit' + i ).prop( "disabled", false );
		$( '#output' + i ).prop( "disabled", false );
		$( '#ratekb' + i ).prop( "disabled", false );
		$( '#allowedcode' + i ).prop( "disabled", false );
		
	}
	
}

var StartIterationvh = 1	
function addrowVH(trid)
{

	var lstvh = this.StartIterationvh++;
	if (lstvh < 2) {
		var row = $("#trheader3").after("<tr id='trv" + lstvh + "'></tr>");
		var kurangbutton = "&nbsp;";
	} else {
		var row = $('#trv' + trid).after("<tr id='trv" + lstvh + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowVH(" + lstvh + ")' />";
	}

	var employeecode = "<select name='employeecode" + lstvh + "' id='employeecode" + lstvh + "' class='lebaractivity'></select>";
	$('#trv' + lstvh).append("<td class='ratakiri'>" + employeecode + "</td>");

	var helpercode = "<select name='helpercode" + lstvh + "' id='helpercode" + lstvh + "' class='lebaractivity'></select>";
	$('#trv' + lstvh).append("<td class='ratakiri'>" + helpercode + "</td>");

	var lokasivh = "<select name='lokasivh" + lstvh + "' id='lokasivh" + lstvh + "' class='ratakiri lebarabsen' onchange='getTCVH(" + lstvh + ")'></select>";
	$('#trv' + lstvh).append("<td class='ratakiri lebarabsen'>" + lokasivh + "</td>");

	var timestart = "<input type='text' name='timestart" + lstvh + "' id='timestart" + lstvh + "' value='' placeholder='hh:mm-hh:mm' autocomplete='off' class='lebartime ratatengah timepicker'>";
	$('#trv' + lstvh).append("<td class='ratakiri'>" + timestart + "</td>");

	var activityivh = "<select name='activityivh" + lstvh + "' id='activityivh" + lstvh + "' class='lebaractivity' onchange='getTCVH(" + lstvh + ")'></select>";
	$('#trv' + lstvh).append("<td class='ratakiri'>" + activityivh + "</td>");

	var tcvh = "<select name='tcvh" + lstvh + "' id='tcvh" + lstvh + "' class='lebartc ratatengah'></select>";
	$('#trv' + lstvh).append("<td class='ratakiri'>" + tcvh + "</td>");

	var kattransport = "<select name='kattransport" + lstvh + "' id='kattransport" + lstvh + "' class='lebaractivity'></select>";
	$('#trv' + lstvh).append("<td class='ratakiri'>" + kattransport + "</td>");

	var trip = "<input type='text' name='trip" + lstvh + "' id='trip" + lstvh + "' value='0' onkeyup='hanyaAngka(this)' class='lebarabsen ratatengah'>";
	$('#trv' + lstvh).append("<td class='ratakiri'>" + trip + "</td>");

	var jamkerja = "<input type='text' name='jamkerja" + lstvh + "' id='jamkerja" + lstvh + "' value='0' onkeyup='hanyaAngka(this)' class='lebarabsen ratatengah'>";
	$('#trv' + lstvh).append("<td class='ratakiri'>" + jamkerja + "</td>");

	var kmhm = "<input type='text' name='kmhm" + lstvh + "' id='kmhm" + lstvh + "' value='0' onkeyup='hanyaAngka(this)' class='lebarabsen ratatengah'>";
	$('#trv' + lstvh).append("<td class='ratakiri'>" + kmhm + "</td>");

	var outputvh = "<input type='text' name='outputvh" + lstvh + "' id='outputvh" + lstvh + "' value='0' onkeyup='hanyaAngka(this)' class='lebarabsen ratatengah'>";
	$('#trv' + lstvh).append("<td class='ratakiri'>" + outputvh + "</td>");

	var remarkvh = "<textarea name='remarkvh" + lstvh + "' id='remarkvh" + lstvh + "' rows='3' cols='40' class='lebaractivity'></textarea>";
	$('#trv' + lstvh).append("<td class='ratakiri'>" + remarkvh + "</td>");

	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowVH(" + lstvh + ")' />";
	$('#trv' + lstvh).append("<td>" + tambahbutton + "</td>");

	$('#trv' + lstvh).append("<td>" + kurangbutton + "</td>");

	$('#jumldatavh').val(lstvh);

	showEmployeeVH("employeecode" + lstvh);
	showEmployeeVH("helpercode" + lstvh);
	showFCBAVH("lokasivh" + lstvh);
	showActivityVH("activityivh" + lstvh);
	showkategoriTR("kattransport" + lstvh);

}


function deleterowVH(i) {
	Spry.Widget.Utils.destroyWidgets('employeecode' + i);
	Spry.Widget.Utils.destroyWidgets('helpercode' + i);
	Spry.Widget.Utils.destroyWidgets('lokasivh' + i);
	Spry.Widget.Utils.destroyWidgets('timestart' + i);
	Spry.Widget.Utils.destroyWidgets('activityivh' + i);
	Spry.Widget.Utils.destroyWidgets('tcvh' + i);
	Spry.Widget.Utils.destroyWidgets('kattransport' + i);
	Spry.Widget.Utils.destroyWidgets('trip' + i);
	Spry.Widget.Utils.destroyWidgets('jamkerja' + i);
	Spry.Widget.Utils.destroyWidgets('kmhm' + i);
	Spry.Widget.Utils.destroyWidgets('outputvh' + i);
	Spry.Widget.Utils.destroyWidgets('remarkvh' + i);
	 $('#trv' + i).remove();
	 
}


//get employee
function showEmployeeVH(selid) {
	$('#employeecode').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : "bkm_employee_edit.asp",
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}


//get activity kendaraan
function showActivityVH(selid) {
    $('#activityivh').empty();
    $('#' + selid).select2({
        ajax: {
            delay: 1000,
            url: function() {
                var veh = $("#kendaraan").val(); // ← dibaca SAAT request, bukan saat init
                return 'bkm_getactivity.asp?panen=4&veh=' + veh;
            },
            dataType: 'json',
            type: "POST",
            allowClear: true,
            data: function(params) {
                return { keyword: params.term };
            }
        }
    });
}

//get kategori transport
function showkategoriTR(selid) {
	
	$('#kattransport').empty();
	var jmlitem = 4;
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_show_item.asp?jmlitem=' + jmlitem,
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}


//get tcvh
function showtcVH(selid) {
	
	$('#tcvh').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getblok.asp',
			dataType: 'json',
			type: "POST",
			allowClear: true,
			data: function (params) {
					var query = {
									keyword: params.term
								}
					return query;
			}
		}
	});
}


new Spry.Widget.ValidationSelect("kendaraan");
</script>