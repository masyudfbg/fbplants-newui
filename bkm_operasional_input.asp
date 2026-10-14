<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim businessunit,fcbaname,CompanyName,q,sqltotdata,gangcode
businessunit= trim(session("fcba"))
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
		<title>FBPlants - BKM RAWAT / MILL</title>
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
<script type="text/javascript" src="/js/fbp-shell.js?v=20261014a"></script>
<% end if %>
</head>

<style>
	.posisikirijudul{
		position: relative;
		left: 0px;
	}
	.posisikiri{
		position: relative;
		left: -15px;
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
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : BKM RAWAT/MILL</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry" method="post" action="bkm_operasional_proses.asp">
	
	<input type="hidden" name="jumldata" id="jumldata" value="0" />
	<input type="hidden" name="jumldataitem" id="jumldataitem" value="1" />
	<input type="hidden" name="proses" id="proses" value="1" />
	<table class="data posisikiri" border="0" width="90%">			
		<tr>
			<td align="left" >&nbsp;&nbsp;&nbsp;Tanggal</td>
			<td align="left">
				<input type="text" class="datepicker" maxlength="15" name="tglbkm" readonly id="tglbkm" size="15" value="<%=textDate(date)%>" title="Tanggal BKM" onchange="get_Allowentry();">
				<!--<input type="text" class="datepicker" maxlength="15" name="tglbkm" readonly id="tglbkm" size="15" value="<%=textDate(date)%>" title="Tanggal BKM" onchange="get_Allowentry();">-->
			</td>
			<td align="left" >No.RKB</td>
			<td align="left">
				<select name="rkbno" id="rkbno" class="lebarsection" onchange="showSection();">
				</select>
			</td>
			
			<td align="left" >Section Name</td>
			<td align="left">
				<select name="sectionname" id="sectionname" class="lebarsection">
				</select>
			</td>
			
			<td align="left" >Gang</td>
			<td align="left">
				<select name="gangcode" id="gangcode" class="lebarsection" onchange="getEmployee();">
				</select>
			</td>
			
			
		</tr>
		<tr>
			
			<td align="left" >Mandor I</td>
			<td align="left">
				<select name="supervisi1" id="supervisi1" class="lebarsection" >
				</select>
			</td>
			<td align="left" >Mandor</td>
			<td align="left">
				<select name="supervisi2" id="supervisi2" class="lebarsection" >
				</select>
			</td>
			
			<td align="left" >Kerani Panen</td>
			<td align="left">
				<select name="supervisi3" id="supervisi3" class="lebarsection" >
				</select>
			</td>
			<!--<td align="left" >Clasification</td>
			<td align="left">
				
				<select name="clasification" id="clasification" class="lebarsection">
				</select>
			</td>-->
			
			<td align="left" >Kerani Transport</td>
			<td align="left">
				<select name="supervisi4" id="supervisi4" class="lebarsection" >
				</select>
			</td>
		</tr>
	</table>
	
	<br>
		
	<table align=center border="0" class='data posisikiri' cellpadding="0" cellspacing="0" width='1100' id='tableall'>
	<input type='hidden' value='1' name='dataunique' id='dataunique'>
		<tr>
			<td colspan=>
				<div class="TabbedPanels" id="tp1" align='center'>
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">BKM UPAH</li>
						<li class="TabbedPanelsTab">BKM BAHAN</li>
						
					</ul>

					<div class="TabbedPanelsContentGroup">

						<div class="TabbedPanelsContent">
							<table align=center border="0" width="1200" class='data' id='table2'>
								<tr id="trheader2">
									
									<th>Employee</th>
									<th>TimeIn</th>
									<th>TimeOut</th>
									<th>Lokasi</th>
									<th>No.RKB</th>
									<th>Absen</th>
									<th>Activity</th>
									<th>TC</th>
									<th>HK</th>
									<th>Unit</th>
									<th>Uom</th>
									<th>Output(ha)</th>
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

								<tr>
									<td colspan='16' align=center><input Type="Submit" id="SIMPAN_UPAH" name="simpan" value="SIMPAN UPAH"  /> </td>
									
								</tr>
								
							</table>
						</div>

						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="datadetail">
								<table align=center border="0"  width="850" class='data' id='table1'>
									<tr id="trheader1">
										<th>Activity</th>
										<th>TC</th>
										<th>Item</th>
										<th>Uom</th>
										<th>Qty</th>
										<th colspan="2">&nbsp;</th>
									</tr>


										<tr id='trb'>
											<tbody id="bodytablebrg">
											</tbody>
											
										</tr>

										<tr>
											<td colspan='5' align=center><input Type="Submit" id="SIMPAN_BAHAN" name="simpan" value="SIMPAN BAHAN"  /></td>
											
										</tr>
									
								</table>
							</div>
							<!-- ============================================================================================================================= -->
						</div>


					</div>
				</div>
			</td>
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
	

	// $('#SIMPAN_BAHAN').click(function(){
	// 	e.preventDefault(); // cegah submit form dulu

    //     let isValid = true;
		
	// 	$("input[name='qtyitem[]']").each(function () {
    //         let val = parseInt($(this).val(), 10) || 0;
    //         if (val === 0) {
    //             alert("Qty tidak boleh 0!");
    //             $(this).focus();
    //             isValid = false;
    //             return false; // stop loop
    //         }
    //     });
	// });

function hanyaAngka(evt) {
	evt.value = evt.value.replace(/[^0-9,.-]/g,"");
}

$('#sectionname').select2();

$(function() {	
	$("#tglbkm").datepicker({
		dateFormat: "mm/dd/yy"
	});
	var section = $('#sectioninput').val();
	getNoRKB();
	//showSection();
	showClasification();
	showSupervisi1();
	showSupervisi2();
	showSupervisi3();
	showSupervisi4();
	addrowBahan(1);

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

function get_Allowentry()
{
	var tglbkm = $('#tglbkm').val();
	if (tglbkm != "") {
		$.post("bkm_allow_entry.asp?tglbkm=" + tglbkm, function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var closingStatus = json.results[0].msg;
			if(closingStatus=="FALSE" || closingStatus=="TRUE"){
				$('#SIMPAN_UPAH').show();
				$('#SIMPAN_BAHAN').show();
				getNoRKB();
				//window.open('bkm_operasional_edit.Asp?bkmno=' + bkmno);
			}else{
				alert("PERIODE SUDAH DI CLOSE !!!");
				$('#SIMPAN_UPAH').hide();
				$('#SIMPAN_BAHAN').hide();
				//break;
			}
			
			
						
		});
		
	} 
}
//GET RKB
function getNoRKB()
{
	var tglbkm = $('#tglbkm').val();
	var lokasi = '<%= Session("fcba") %>';
	$('#rkbno').empty();
	$('#rkbno').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getrkb.asp?tglbkm=' + tglbkm + '&fcba=' + lokasi + '&id=' + 1+'&doc='+501,
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

//get RKB Lokasi upah
function getNoRKBFCBA(selid) {
	var tglbkm = $('#tglbkm').val();
	var lokasi = $('#lokasi' + selid).val();
	$('#rkbnofcba' + selid).empty();
	$('#rkbnofcba' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getrkb.asp?tglbkm=' + tglbkm + '&fcba=' + lokasi + '&id=' + 1+'&doc='+501,
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


//GET EMPLOYEE
function getEmployee()
{
	var gangcode = $("#gangcode").val();
	var tglbkm = $('#tglbkm').val();
	if (gangcode != "") {
		$.post("bkm_employee.asp?gangcode=" + gangcode + "&tglbkm=" + tglbkm + "&employee=" + 100, function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var jumlview  = json.results.length;
			if(jumlview>0){
				var bkmno = json.results[0].bkmno;
				var postingflg = json.results[0].postingflg;
				//if(postingflg==0){
				//	window.open('bkm_operasional_edit.Asp?bkmno=' + bkmno + '&postingflg=' + postingflg);
				window.close();
				window.open('bkm_operasional_edit.Asp?bkmno=' + bkmno);
				//}
			}else{
				
					getEmployeeview();
				
			}
			
			
						
		});
		
	} 
}


//get section
function showSection() {
	var rkbno = $('#rkbno').val();
	
	$('#sectionname').empty();
	$('#sectionname').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getsection.asp?section=' + 9 +'&rkbno='+rkbno,
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

$('#sectionname').on('change', function(){
	let val = $(this).val();
	$('#sectioninput').val(val);
	showGang(val);
	showItem('ITEMCODE1', val);
});

//get GANGCODE 
function showGang(val) {
	$('#gangcode').empty();
	$('#gangcode').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_gangcode.asp?sectionname=' + val + '&gancode=' + 1,
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

//get clasification 
function showClasification() {
	$('#clasification').empty();
	$('#clasification').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_clasification.asp',
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

//get supervisi 1
function showSupervisi1() {
	
	$('#supervisi1').select2({
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

//get supervisi 2
function showSupervisi2() {
	
	$('#supervisi2').select2({
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

//get supervisi 3
function showSupervisi3() {
	
	$('#supervisi3').select2({
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

//get supervisi 1
function showSupervisi4() {
	
	$('#supervisi4').select2({
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

//get activity
function showActivity(selid) {
	var rkbno = $('#rkbnofcba' + selid).val();
	$('#activity').empty();
	$('#activity' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp?panen='+9+'&rkbno='+rkbno,
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

//getEmployee
function getEmployeeview() {
	
	var gangcode = $("#gangcode").val();
	var tglbkm = $('#tglbkm').val();
	var rkbno = $('#rkbno').val();
	
	if (gangcode != "") {
		$.post("bkm_employee.asp?gangcode=" + gangcode + "&tglbkm=" + tglbkm + "&employee=" + 1, function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var jumlview  = json.results.length;
			var no = 1;
			for (var i = 0; i < jumlview; i++) {
				var fccode = json.results[i].fccode;
				var fcname = json.results[i].fcname;
				var presentgrade = json.results[i].presentgrade;
				var fcba 		= json.results[i].fcbaname;
				var dttimein 	= json.results[i].timeinx;
				var dttimeout 	= json.results[i].timeoutx;
				var datatimein 	= json.results[i].timein;
				var datatimeout = json.results[i].timeout;
				var ttlmandays 	= json.results[i].ttlmandays;
				//var hkhi 		= json.results[i].mandayshi;
						
				
				var docketlist = '';
					docketlist += '<tr id="tru'+no+'">';
					docketlist += '<td align="left"  class="ratakiri lebaractivity "><select class="ratakiri lebaractivity" name="fccode' + no + '" id="fccode' + no + '" onchange="getFinger(' + no + ');getHKhi(' + no + ');"><option value=' + fccode + '>' + fcname + '</option></select><input class="ratatengah" type="hidden" name="statusemp' + no + '" id="statusemp' + no + '" value="' + presentgrade + '"/></td>';
					docketlist += '<td align="center">' + datatimein + '<input type="hidden" class="ratatengah lebarhk" name="timeinx' + no + '" id="timeinx' + no + '" value="' + dttimein + '"/></td>';
					docketlist += '<td align="center">' + datatimeout + '<input type="hidden" class="ratatengah lebarhk" name="timeoutx' + no + '" id="timeoutx' + no + '" value="' + dttimeout + '"/></td>';
					docketlist += '<td align="left" class="ratakiri lebarabsen"><select class="ratakiri lebarlokasi" name="lokasi' + no + '" id="lokasi' + no + '" onchange="getNoRKBFCBA(' + no + ')"><option value=' + fcba + '>' + fcba + '</option></select></td>';
					docketlist += '<td align="left" class="ratakiri lebaractivity"><select class="ratakiri lebaractivity" name="rkbnofcba' + no + '" id="rkbnofcba' + no + '" onchange="showActivity(' + no + ')"><option value=' + rkbno + '>' + rkbno + '</option></select></td>';
					docketlist += '<td align="left" class="ratakiri lebarabsen"><select class="ratakiri lebarabsen" name="absen' + no + '" id="absen' + no + '" onchange="getDisable(' + no + ')"/></select></td>';
					docketlist += '<td align="center" class="ratakiri lebaractivity"><select class="ratakiri lebaractivity" name="activity' + no + '" id="activity' + no + '" onchange="getTC(' + no + ');getuom(' + no + ');"/></select></td>';
					docketlist += '<td align="center"><select class="ratatengah lebartc" name="tcupah' + no + '" id="tcupah' + no + '"/></select></td>';		
					docketlist += '<td align="center"><input class="ratatengah lebarhk" type="text" onkeyup="hanyaAngka(this);getValidasiHK(' + no + ');" name="hk' + no + '" id="hk' + no + '" value="0"/><input class="ratatengah" type="hidden" onkeyup="hanyaAngka(this)" name="ttlhkkhl' + no + '" id="ttlhkkhl' + no + '" value="' + ttlmandays + '"/></td>';
					docketlist += '<td align="center"><input class="ratatengah lebarabsen" type="text" onkeyup="hanyaAngka(this)" name="unit' + no + '" id="unit' + no + '" value="0"/></td>';
					docketlist += '<td align="center"><input readonly class="ratatengah lebarabsen" type="text" name="uom_unit' + no + '" id="uom_unit' + no + '" value="HA" /></td>';
					
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
				showActivity(no);
				showAllowed("allowedcode" + no);
				showTypeOv("typeovrtime" + no);

				
				new Spry.Widget.ValidationSelect("fccode" + no);
				new Spry.Widget.ValidationTextField("statusemp" + no);
				new Spry.Widget.ValidationSelect("lokasi" + no);
				new Spry.Widget.ValidationSelect("absen" + no);
				new Spry.Widget.ValidationSelect("activity" + no);
				new Spry.Widget.ValidationSelect("tcupah" + no);
				new Spry.Widget.ValidationTextField("hk" + no);
				//new Spry.Widget.ValidationTextField("ttlhkkhl" + no);

				
				
				no++;
				$("#jumldata").val(jumlview);
			}
			
		});
		
	} else {
		$('#bodytable').empty();
	}
	
}

function getValidasiHK(i)
{
	var tglbkm = $('#tglbkm').val();
	var thnblntgl  = tglbkm.split('/');
	var blnview 	= thnblntgl[0];
	var tglview 	= thnblntgl[1];
	var thnview 	= thnblntgl[2];
	var thnblntglview = thnview + blnview + tglview;
	//alert(thnblntglview);
	var statusemp = $('#statusemp' + i).val();
	var ttlhkkhl = $('#ttlhkkhl' + i).val();
	var hk = $('#hk' + i).val();
	if(statusemp=="KHL" & parseFloat(ttlhkkhl)> 20 & thnblntglview>20230126){
		alert("KARYAWAN DENGAN STATUS KHL TIDAK BOLEH MELEBIHI DARI 20 HK");
		$('#hk' + i).val(0);
	}else if(hk>1){
		alert("INPUTAN HK TIDAK BOLEH LEBIH DARI 1");
		$('#hk' + i).val(0);
	}
	
}


function getTC(i)
{
    var rkbno    = $('#rkbnofcba' + i).val();
	var activity = $('#activity' + i).val();
	var lokasi   = $('#lokasi' + i).val();
	var tglbkm   = $('#tglbkm').val();
	$('#tcupah' + i).empty();
	$('#tcupah' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getblok.asp?blok='+9+'&activity=' + activity +'&lokasi=' + lokasi +'&rkbno=' + rkbno + '&tglbkm=' + tglbkm,
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

//ambil data satuan
function getuom(i)
{
	var activity = $('#activity' + i).val();
	if (activity != "") {
		$.post("bkm_getactivity.asp?activity=" + activity +"&panen=" + 10, function(data) {
			$('#uom_unit' + i).empty();
			var json  = JSON.parse(data);
			var uom = json.results[0].uom_unit
			$('#uom_unit' + i).val(uom);
						
		});
		
	}
}


function getTCItem(i)
{
    var rkbno = $('#rkbno').val();
	var tglbkm = $('#tglbkm').val();
	var activityitem = $('#activityitem' + i).val();
	var lokasi = '<%= Session("fcba") %>';

	$('#tcitem' + i).empty();
	$('#tcitem' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getblok.asp?blok='+9+'&activity=' + activityitem +'&lokasi=' + lokasi +'&rkbno=' + rkbno ,
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


//get item
function showItem(selid, sections) {
	var tglbkm = $('#tglbkm').val();
	$('#ITEMCODE').empty();
	var jmlitem = 0;
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : "bkm_show_item.asp?jmlitem=" + jmlitem + "&tglbkm=" + tglbkm + "&section=" + sections,
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

//ambil data satuan
function getSatuan(i)
{
	var itemcode = $('#ITEMCODE' + i).val();
	if (itemcode != "") {
		$.post("bkm_employee.asp?itemcode=" + itemcode +"&satuan=" + 1, function(data) {
			$('#uom' + i).empty();
			var json  = JSON.parse(data);
			var uom = json.results[0].unitofmeasurement
			$('#uom' + i).val(uom);
						
		});
		
	}
}

//get activity item
function showActivityitem(selid) {
	var rkbno = $('#rkbno').val();
	$('#activityitem').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp?panen='+9+'&rkbno='+rkbno,
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


//var StartIteration=1;
function addrowUpah(trid,nourut) {
	
     var fcba = '<%= Session("fcba") %>';
	 var rkbno = $('#rkbno').val();
	 var lastrow	= parseInt(nourut)+1;
	//  var rowx       = $('#jumldata').val();
	//  var lastrow	= parseInt(rowx)+1;
	
	if (lastrow < 2) {
		var row = $("#trheader2").after("<tr id='tru" + lastrow + "'></tr>");
		var kurangbutton = "&nbsp;";
	} else {
		var row = $('#tru' + trid).after("<tr id='tru" + lastrow + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowUpah(" + lastrow + ");' />";
	}

	//var lastrow = this.StartIteration++;
	
	
	var fccode = "<select name='fccode" + lastrow + "' id='fccode" + lastrow + "' class='ratakiri lebaractivity' onchange='getFinger(" + lastrow + ");'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri'>" + fccode + "</td>");

	var timein = "<span name='timein" + lastrow + "' id='timein" + lastrow + "' class='ratatengah lebarrate'></span><input type='hidden' name='timeinx" + lastrow + "' id='timeinx" + lastrow + "'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + timein + "</td>");

	var timeout = "<span name='timeout" + lastrow + "' id='timeout" + lastrow + "' class='ratatengah lebarrate'></span><input type='hidden' name='timeoutx" + lastrow + "' id='timeoutx" + lastrow + "'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + timeout + "</td>");

	var lokasi = "<select name='lokasi" + lastrow + "' id='lokasi" + lastrow + "' class='ratakiri lebarlokasi'><option value='" + fcba + "'>" + fcba + "</option></select>";
	$('#tru' + lastrow).append("<td class='ratakiri lebarlokasi'>" + lokasi + "</td>");

	var rkbnofcba = "<select name='rkbnofcba" + lastrow + "' id='rkbnofcba" + lastrow + "' class='ratakiri lebaractivity' onchange='showActivity(" + lastrow + ");'><option value='" + rkbno + "'>" + rkbno + "</option></select>";
	$('#tru' + lastrow).append("<td class='ratakiri lebaractivity'>" + rkbnofcba + "</td>");

	var absen = "<select name='absen" + lastrow + "' id='absen" + lastrow + "' class='ratakiri lebarabsen' onchange='getDisable(" + lastrow + ");'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri lebarabsen'>" + absen + "</td>");

	var activity = "<select name='activity" + lastrow + "' id='activity" + lastrow + "' class='ratakiri lebaractivity' onchange='getTC(" + lastrow + ");getuom(" + lastrow + ");getDisable(" + lastrow + ");'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri lebaractivity'>" + activity + "</td>");

	var blok = "<select type='text' name='tcupah" + lastrow + "' id='tcupah" + lastrow + "' value='0' class='ratatengah lebartc'></select>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + blok + "</td>");

	var hk = "<input type='text' name='hk" + lastrow + "' id='hk" + lastrow + "' onkeyup='hanyaAngka(this);getHKhinew(" + lastrow + ")' value='0' class='ratatengah lebarhk'><input class='ratatengah' type='hidden' onkeyup='hanyaAngka(this)' name='ttlhkkhl" + lastrow + "' id='ttlhkkhl" + lastrow + "' value='0'/>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + hk + "</td>");

	var unit = "<input type='text' name='unit" + lastrow + "' id='unit" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + unit + "</td>");

	var uom_unit = "<input type='text' name='uom_unit" + lastrow + "' id='uom_unit" + lastrow + "'  value='HA' readonly class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + uom_unit + "</td>");


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


	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowUpah(" + lastrow + "," + lastrow + ");' />";
	$('#tru' + lastrow).append("<td>" + tambahbutton + "</td>");

	$('#tru' + lastrow).append("<td>" + kurangbutton + "</td>");

	$('#jumldata').val(lastrow);
	

	showEmployee("fccode" + lastrow);
    showFCBA("lokasi" + lastrow);
	showAbsen("absen" + lastrow);
	showActivity(lastrow);
	showAllowed("allowedcode" + lastrow);
	showTypeOv("typeovrtime" + lastrow);
	//getjmlRow();

	new Spry.Widget.ValidationSelect("fccode" + lastrow);
	new Spry.Widget.ValidationSelect("absen" + lastrow);
	new Spry.Widget.ValidationSelect("activity" + lastrow);
	new Spry.Widget.ValidationSelect("tcupah" + lastrow);
	new Spry.Widget.ValidationTextField("hk" + lastrow);

}

function deleterowUpah(i) {
	Spry.Widget.Utils.destroyWidgets('fccode' + i);
	// Spry.Widget.Utils.destroyWidgets('timein' + i);
	// Spry.Widget.Utils.destroyWidgets('timeout' + i);
	Spry.Widget.Utils.destroyWidgets('timeinx' + i);
	Spry.Widget.Utils.destroyWidgets('timeoutx' + i);
	Spry.Widget.Utils.destroyWidgets('lokasi' + i);
	Spry.Widget.Utils.destroyWidgets('rkbnofcba' + i);
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
	// getjmlRow();
	
}

//show time from finger employee
function getFinger(i)
{
	var tglbkm = $('#tglbkm').val();
	var fccode = $('#fccode' + i).val();
	
	if (fccode != "") {
		$.post("bkm_employee_edit.asp?emp=" + 1 + "&tglbkm=" + tglbkm + "&fccode=" + fccode, function(data) {
			var json  = JSON.parse(data);
			//var jumlview  = json.results.length;
			var timein = json.results[0].timein;
			var timeinx = json.results[0].timeinx;
			var timeout = json.results[0].timeout;
			var timeoutx = json.results[0].timeoutx;
				$('#timein' + i).html(timein);
				$('#timeout' + i).html(timeout);
				$('#timeinx' + i).val(timeinx);
				$('#timeoutx' + i).val(timeoutx);
			
			
		});
		
	}

}

var StartIterationitem = 1;
function addrowBahan(trid)
{
	var lstitem = this.StartIterationitem++;
	
	if (lstitem < 2) {
		var row = $("#trheader1").after("<tr id='trb" + lstitem + "'></tr>");
		var kurangbutton = "&nbsp;";
	} else {
		var row = $('#trb' + trid).after("<tr id='trb" + lstitem + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowBahan(" + lstitem + ");' />";
	}

	var activityitem = "<select name='activityitem" + lstitem + "' id='activityitem" + lstitem + "' class='lebaractbarang' onchange='getTCItem(" + lstitem + ")'></select>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + activityitem + "</td>");

	var tcitem = "<select name='tcitem" + lstitem + "' id='tcitem" + lstitem + "' class='lebartc ratatengah'></select>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + tcitem + "</td>");

	var ITEMCODE = "<select name='ITEMCODE" + lstitem + "' id='ITEMCODE" + lstitem + "' onchange='getSatuan(" + lstitem + ")' class='lebaritembarang'></select>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + ITEMCODE + "</td>");

	var uom = "<input type='text' name='uom" + lstitem + "' id='uom" + lstitem + "' value='' readonly class='lebarabsen ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + uom + "</td>");

	var qtyitem = "<input type='text' name='qtyitem" + lstitem + "' id='qtyitem" + lstitem + "' value='' onkeyup='hanyaAngka(this)' class='lebarabsen ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + qtyitem + "</td>");

	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowBahan(" + lstitem + ");' />";
	$('#trb' + lstitem).append("<td>" + tambahbutton + "</td>");

	$('#trb' + lstitem).append("<td>" + kurangbutton + "</td>");

	$('#jumldataitem').val(lstitem);

	showItem("ITEMCODE" + lstitem, $('#select2-sectionname-container').attr('title'));

	$('#qtyitem'+lstitem).on('keyup', function() {
		let val = parseFloat($(this).val().replace(/,/g, ''));
		
		// ambil stok dari text
		let text = $("#select2-ITEMCODE" + lstitem + "-container").attr('title');
		let parts = text.split("STOK :");

		let stok = 0;
		if (parts.length > 1) {
			stok = parseFloat(parts[1].trim().replace(/,/g, ''));
		}
		// cek apakah input lebih besar dari stok
		if (val > stok) {
			alert("Jumlah melebihi stok (" + stok + ")");
			$(this).val(0); // optional: reset ke nilai stok
		}
	});

	showActivityitem("activityitem" + lstitem);
	

	// new Spry.Widget.ValidationSelect("activityitem" + lstitem);
	// new Spry.Widget.ValidationSelect("tcitem" + lstitem);
	// new Spry.Widget.ValidationSelect("ITEMCODE" + lstitem);
	// new Spry.Widget.ValidationTextField("qtyitem" + lstitem);
	
}

//hapus row bahan
function deleterowBahan(i) {
	Spry.Widget.Utils.destroyWidgets('ITEMCODE' + i);
	Spry.Widget.Utils.destroyWidgets('uom' + i);
	Spry.Widget.Utils.destroyWidgets('activityitem' + i);
	Spry.Widget.Utils.destroyWidgets('tcitem' + i);
	Spry.Widget.Utils.destroyWidgets('qtyitem' + i);
	 $('#trb' + i).remove();
	// getjmlRowbhn();
}

//hitung jumlah row pada tabel upah
function getjmlRow()
{
	
	var oRows = document.getElementById('table2').getElementsByTagName('tr');
	var iRowCount = oRows.length;
	$('#jumldata').val(iRowCount);

	// var jmlrow = $('#jumldata').val();
	// jmlrow     = parseInt(jmlrow) + 1;
	// $('#jumldata').val(jmlrow);
	
}

//kurangi jumlah row pada tabel upah
function getdeljmlRow()
{
	
	var jmlrow = $('#jumldata').val();
	jmlrow     = parseInt(jmlrow) - 1;
	$('#jumldata').val(jmlrow);
	
}

//hitung jumlah row pada tabel bahan
function getjmlRowbhn()
{
	
	var oRows = document.getElementById('table1').getElementsByTagName('tr');
	var iRowCount = oRows.length;
	$('#jumldataitem').val(iRowCount);
	
}

//function disabled input()
function getDisable(i)
{
	var absen = $('#absen' + i).val();
	if (absen=="KB" || absen=="KH"){
		//disbale
		$( '#hk' + i ).prop( "disabled", true );
		$( '#hk' + i ).val(0);
		//$( '#ratepremi' + i ).prop( "disabled", true );
		//$( '#allowedcode' + i ).prop( "disabled", true );
		
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

function getHKhinew(i)
{
	var emp 	= $('#fccode' + i).val();
	var tglbkm 	= $('#tglbkm').val();
	var hk 		= $('#hk' + i).val();

	if (emp != "") {
		$.post("bkm_employee.asp?emp=" + emp + "&employee=" + 300 + "&tglbkm=" + tglbkm, function(data) {
			$('#hk' + i).empty();
			var json  = JSON.parse(data);
			var hkhi  = json.results[0].mandayshi;
			var ttlhk = parseFloat(hkhi)+parseFloat(hk);
			
			if((ttlhk)>1){
				alert("TOTAL HK PEKERJA INI SUDAH LEBIH DARI 1 HK, SILAHKAN CEK KEMBALI INPUTAN ANDA..");
				 $('#hk' + i).val("0");
				//  $('#fccode' + i).empty();
				//  $('#lokasi' + i).empty();
				//  $('#absen' + i).empty();
				//  $('#activity' + i).empty();
				//  $('#tcupah' + i).empty();
				//  $('#timein' + i).empty();
				//  $('#timeout' + i).empty();
				//  $('#timeinx' + i).empty();
				//  $('#timeoutx' + i).empty();
				 
			}
						
		});
		
	}

	if(hk>1){
		alert("INPUTAN HK TIDAK BOLEH LEBIH DARI 1 HK");
		$('#hk' + i).val(0);
	}
}



//new Spry.Widget.ValidationTextField("DATECREATED", "none");
//new Spry.Widget.ValidationSelect("clasification");
</script>