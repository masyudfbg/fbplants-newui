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
		<title>FBPlants - BKM PANEN</title>
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
	.lebarkgtbs{
		width: 50px;
	}
	.lebaractivity{
		width: 95px;
	}
	.lebaremp{
		width: 120px;
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
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : BKM PANEN</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry" method="post" action="bkm_panen_proses.asp">
	
	<input type="hidden" name="jumldata" id="jumldata" value="0" />
	<input type="hidden" name="jumldataitem" id="jumldataitem" value="1" />
	<input type="hidden" name="jumldatadenda" id="jumldatadenda" value="1" />
	<input type="hidden" name="proses" id="proses" value="1" />
	<table class="data posisikiri" border="0" width="90%">			
		<tr>
			<td align="left" >&nbsp;&nbsp;&nbsp;Tanggal</td>
			<td align="left">
				<input type="text" class="datepicker" maxlength="15" name="tglbkm" readonly id="tglbkm" size="15" value="<%=textDate(date)%>" title="Tanggal BKM" onchange="get_Allowentry();">
			</td>
			<td align="left" >No.RKB</td>
			<td align="left">
				<select name="rkbno" id="rkbno" class="lebarsection" onchange="showSection();">
				</select>
			</td>
			<td align="left" >Section Name</td>
			<td align="left">
				<select name="sectionname" id="sectionname" class="lebarsection" onchange="showGang();">
				</select>
			</td>
			<td align="left" >&nbsp;&nbsp;&nbsp;Gang</td>
			<td align="left">
				<select name="gangcode" id="gangcode" class="lebargangcode" onchange="getEmployee();">
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
			
			<td align="left" >Kerani Transport</td>
			<td align="left">
				<select name="supervisi4" id="supervisi4" class="lebarsection" >
				</select>
			</td>
			<td align="left" >Hujan</td>
			<td align="left">
				<input type="checkbox" name="hujan" id="hujan" class="" >
				</input>
			</td>
		</tr>
	</table>
	
	<br>
		
	<table align=center border="0" class='data posisikiri' width='1100' id='tableall'>
		<tr>
			<td colspan=>
				<div class="TabbedPanels" id="tp1" align='center'>
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">BKM UPAH</li>
						<li class="TabbedPanelsTab">BKM MUTU</li>
						<!--<li class="TabbedPanelsTab">BKM DENDA</li>-->
						
					</ul>

					<div class="TabbedPanelsContentGroup">

						<div class="TabbedPanelsContent">
							<table align=center border="0" width="1000" class='data' id='table2'>
								<tr id="trheader2">
									
									<th>Employee</th>
									<th>TimeIn</th>
									<th>TimeOut</th>
									<th>Lokasi</th>
									<th>No. RKB</th>
									<th>Absen</th>
									<th>Activity</th>
									<th>TC</th>
									<th>HK</th>
									<th>JJG</th>
									<th>JJG UN</th>
									<th>JJG BA</th>
									<th>BJR</th>
									<th>Brdln</th>
									<th>KG TBS</th>
									<th>Output</th>
									<th>Rate</th>
									<th>Allowed Code</th>
									<th>Rate Allowed</th>
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
								<table align=center border="0" width="850" class='data' id='table1'>
									<tr id="trheader1">
										
											<th>Blok</th>
											<th>B. Mateng</th>
											<th>B. LMateng</th>
											<th>B. KMateng</th>
											<th>B. Mentah</th>
											<th>B. Busuk</th>
											<th colspan="2">&nbsp;</th>
										</tr>

									<tr id='trb'>
										<tbody id="bodytablebrg">
										</tbody>
										
									</tr>
									
								</table>
							</div>
							<!-- ============================================================================================================================= -->
						</div>

						<!--<div class="TabbedPanelsContent">
							
							<div id="detaildenda">
								<table align=center border="0" width="950" class='data' id='table3'>
									<tr >
										
											<th rowspan="2">Employee</th>
											<th colspan="7">Kualitas TBS</th>
											<th colspan="7">Kualitas Ancak</th>
											<th rowspan="2" colspan="2">&nbsp;</th>
									</tr>
									<tr id="trheader3">
										
										<th>(A)</th>
										<th>(BA)</th>
										<th>(TP)</th>
										<th>(BK)</th>
										<th>(O)</th>
										<th>(AB)</th>
										<th>(E)</th>

										<th>(BT)</th>
										<th>(PS)</th>
										<th>(SP)</th>
										<th>(BRD)</th>
										<th>(OP)</th>
										<th>(BM)</th>
										<th>(BDT)</th>
									</tr>

									<tr id='trd'>
										<tbody id="bodytablebdenda">
										</tbody>
										
									</tr>
									
								</table>
							</div>
							
						</div> -->

						
					</div>
				</div>
			</td>
		</tr>
		<tr>
			<td colspan='' align=center><input Type="Submit"  id="SIMPAN_UPAH" name="simpan" value="Simpan"></td>
			
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
	getNoRKB();
	//showSection();
	showSupervisi1();
	showSupervisi2();
	showSupervisi3();
	showSupervisi4();
	addrowMutu(1);
	showBlokmutu();
	//addrowDenda(1)
	
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

//get GANGCODE 
function showGang() {
	var sectionname = $("#sectionname").val();
	$('#gangcode').empty();
	$('#gangcode').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_gangcode.asp?sectionname=' + sectionname + '&gancode=' + 2,
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
	var rkbno = $('#rkbno').val();
	$('#activity').empty();
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

function get_Allowentry()
{
	var gangcode = $("#gangcode").val();
	var tglbkm = $('#tglbkm').val();
	if (tglbkm != "") {
		$.post("bkm_allow_entry.asp?tglbkm=" + tglbkm, function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var closingStatus = json.results[0].msg;
			if(closingStatus=="FALSE" || closingStatus=="TRUE"){
				$('#SIMPAN_UPAH').show();
				getNoRKB();
				//window.open('bkm_operasional_edit.Asp?bkmno=' + bkmno);
			}else{
				alert("PERIODE SUDAH DI CLOSE !!!");
				$('#SIMPAN_UPAH').hide();
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
			url : 'bkm_getrkb.asp?tglbkm=' + tglbkm + '&id=' + 1+'&doc='+503+'&fcba=' +lokasi ,
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
		$.post("bkm_employee.asp?gangcode=" + gangcode + "&tglbkm=" + tglbkm + "&employee=" + 2, function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var jumlview  = json.results.length;
			var no = 1;
			for (var i = 0; i < jumlview; i++) {
				var fccode = json.results[i].fccode;
				var fcname = json.results[i].fcname;
				var fcba = json.results[i].fcbaname;
				var datatimein 	= json.results[i].timein;
				var datatimeout = json.results[i].timeout;
				var timeinx 	= json.results[i].timeinx;
				var timeoutx 	= json.results[i].timeoutx;
				var presentgrade = json.results[i].presentgrade;
				var ttlmandays 	= json.results[i].ttlmandays;

				if(ttlmandays=="" || ttlmandays=="undefined"){
					ttlmandays = "0";
				}

			
						
				
				var docketlist = '';
					docketlist += '<tr id="tru'+no+'">';
					docketlist += '<td align="left"  class="ratakiri lebaractivity "><select class="ratakiri lebaractivity" name="fccode' + no + '" id="fccode' + no + '" onchange="getFinger(' + no + ');"><option value=' + fccode + '>' + fcname + '</option></select><input class="ratatengah" type="hidden" name="statusemp' + no + '" id="statusemp' + no + '" value="' + presentgrade + '"/></td>';
					docketlist += '<td align="center">' + datatimein + '<input type="hidden" class="ratatengah lebarhk" name="timeinx' + no + '" id="timeinx' + no + '" value="' + timeinx + '"/></td>';
					docketlist += '<td align="center">' + datatimeout + '<input type="hidden" class="ratatengah lebarhk" name="timeoutx' + no + '" id="timeoutx' + no + '" value="' + timeoutx + '"/></td>';
					docketlist += '<td align="left" class="ratakiri lebarabsen"><select class="ratakiri lebarlokasi" name="lokasi' + no + '" id="lokasi' + no + '" onchange="getNoRKBFCBA(' + no + ')"><option value=' + fcba + '>' + fcba + '</option></select></td>';
					docketlist += '<td align="left" class="ratakiri lebaractivity"><select class="ratakiri lebaractivity" name="rkbnofcba' + no + '" id="rkbnofcba' + no + '" onchange="showActivity(' + no + ')"><option value=' + rkbno + '>' + rkbno + '</option></select></td>';
					docketlist += '<td align="left" class="ratakiri lebarabsen"><select class="ratakiri lebarabsen" name="absen' + no + '" id="absen' + no + '" onchange="getDisable(' + no + ')"/></select></td>';
					docketlist += '<td align="center" class="ratakiri lebaractivity"><select class="ratakiri lebaractivity" name="activity' + no + '" id="activity' + no + '" onchange="getTC(' + no + ');"/></select></td>';
					docketlist += '<td align="center"><select class="ratatengah lebartc" name="tcupah' + no + '" id="tcupah' + no + '"/></select></td>';		
					docketlist += '<td align="center"><input class="ratatengah lebarhk" type="text" onkeyup="hanyaAngka(this);getValidasiHK(' + no + ');" name="hk' + no + '" id="hk' + no + '" value="0"/><input class="ratatengah" type="hidden" onkeyup="hanyaAngka(this)" name="ttlhkkhl' + no + '" id="ttlhkkhl' + no + '" value="' + ttlmandays + '"/></td>';
					docketlist += '<td align="center"><input class="ratatengah lebarabsen" type="text" onkeyup="hanyaAngka(this)" name="bunch' + no + '" id="bunch' + no + '" value="0"/></td>';
					docketlist += '<td align="center"><input class="ratatengah lebarabsen" type="text" onkeyup="hanyaAngka(this)" name="jjgu' + no + '" id="jjgu' + no + '" value="0"/></td>';
					docketlist += '<td align="center"><input class="ratatengah lebarabsen" type="text" onkeyup="hanyaAngka(this)" name="jjgab' + no + '" id="jjgab' + no + '" value="0"/></td>';
					docketlist += '<td align="center"><input class="ratatengah lebarabsen" type="text" onkeyup="hanyaAngka(this)" name="abw' + no + '" id="abw' + no + '" value="0"/></td>';
					docketlist += '<td align="center"><input class="ratatengah lebarabsen" type="text" onkeyup="hanyaAngka(this)" name="kgbrondolan' + no + '" id="kgbrondolan' + no + '" value="0"/></td>';
					docketlist += '<td align="center"><input class="ratatengah lebarkgtbs" type="text" onkeyup="hanyaAngka(this)" name="kgtbs' + no + '" id="kgtbs' + no + '" value="0"/></td>';
					docketlist += '<td align="center"><input class="ratatengah lebarabsen" type="text" onkeyup="hanyaAngka(this)" name="output' + no + '" id="output' + no + '" value="0"/></td>';					
					docketlist += '<td align="center"><input class="ratatengah lebarkgtbs" type="text" onkeyup="hanyaAngka(this)" name="ratekb' + no + '" id="ratekb' + no + '" value="0"/></td>';
					docketlist += '<td align="center" class="ratakiri lebaractivity"><select class="ratakiri lebarallow" name="allowedcode' + no + '" id="allowedcode' + no + '"/></select></td>';
					docketlist += '<td align="center"><input class="ratatengah lebarrate" type="text" onkeyup="hanyaAngka(this)" name="ratepremi' + no + '" id="ratepremi' + no + '" value="0"/></td>';	
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
				getNoRKBFCBA("rkbnofcba" + no);

				new Spry.Widget.ValidationSelect("fccode" + no);
				new Spry.Widget.ValidationSelect("absen" + no);
				new Spry.Widget.ValidationSelect("activity" + no);
				new Spry.Widget.ValidationSelect("tcupah" + no);
				new Spry.Widget.ValidationTextField("hk" + no);
				new Spry.Widget.ValidationTextField("bunch" + no);
				
				
				no++;

			}
			
			$("#jumldata").val(jumlview);
			//alert(StartIteration);
			
		});
		
	} else {
		$('#bodytable').empty();
	}
	
}



function getTC(i)
{
	var rkbno = $('#rkbnofcba' + i).val();
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

function showBlokmutu(selid)
{
	
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getblok.asp?blok=' + 1,
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

//function disabled input()
function getDisable(i)
{
	var absen = $('#absen' + i).val();
	if (absen=="KB" || absen=="KH"){
		//disbale
		$( '#hk' + i ).prop( "disabled", true );
		$( '#hk' + i ).val(0);
	}else{
	    $( '#hk' + i ).prop( "disabled", false );
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

	var statusemp = $('#statusemp' + i).val();
	var ttlhkkhl = $('#ttlhkkhl' + i).val();
	var hk = $('#hk' + i).val();
	if(statusemp=="KHL" & parseFloat(ttlhkkhl)>=10 & thnblntglview>20230123){
		alert("KARYAWAN DENGAN STATUS KHL TIDAK BOLEH MELEBIHI DARI 20 HK");
		$('#hk' + i).val(0);
	}else if(hk>1){
		alert("INPUTAN HK TIDAK BOLEH LEBIH DARI 1");
		$('#hk' + i).val(0);
	}
	
}

//var StartIteration=1;
function addrowUpah(trid,nourut) {
	
	 var fcba = '<%= Session("fcba") %>';
	 var lastrow	= parseInt(nourut)+1;
	// alert(lastrow);
	if (lastrow < 2) {
		var row = $("#trheader2").after("<tr id='tru" + lastrow + "'></tr>");
		var kurangbutton = "&nbsp;";
	} else {
		var row = $('#tru' + trid).after("<tr id='tru" + lastrow + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowUpah(" + lastrow + ");' />";
	}
	
	
	var fccode = "<select name='fccode" + lastrow + "' id='fccode" + lastrow + "' class='ratakiri lebaractivity' onchange='getFinger(" + lastrow + ");'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri'>" + fccode + "</td>");

	var timein = "<span name='timein" + lastrow + "' id='timein" + lastrow + "' class='ratatengah lebarrate'></span><input type='hidden' name='timeinx" + lastrow + "' id='timeinx" + lastrow + "'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + timein + "</td>");

	var timeout = "<span name='timeout" + lastrow + "' id='timeout" + lastrow + "' class='ratatengah lebarrate'></span><input type='hidden' name='timeoutx" + lastrow + "' id='timeoutx" + lastrow + "'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + timeout + "</td>");

	var lokasi = "<select name='lokasi" + lastrow + "' id='lokasi" + lastrow + "' class='ratakiri lebarlokasi' onchange='getNoRKBFCBA(" + lastrow + ");'><option value='" + fcba + "'>" + fcba + "</option></select>";
	$('#tru' + lastrow).append("<td class='ratakiri lebarabsen'>" + lokasi + "</td>");

	var rkbnofcba = "<select name='rkbnofcba" + lastrow + "' id='rkbnofcba" + lastrow + "' class='ratakiri lebaractivity' onchange='showActivity(" + lastrow + ");'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri lebaractivity'>" + rkbnofcba + "</td>");

	var absen = "<select name='absen" + lastrow + "' id='absen" + lastrow + "' class='ratakiri lebarabsen' onchange='getDisable(" + lastrow + ");'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri lebarabsen'>" + absen + "</td>");

	var activity = "<select name='activity" + lastrow + "' id='activity" + lastrow + "' class='ratakiri lebaractivity' onchange='getTC(" + lastrow + ");'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri'>" + activity + "</td>");

	var blok = "<select type='text' name='tcupah" + lastrow + "' id='tcupah" + lastrow + "' value='0' class='ratatengah lebartc'></select>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + blok + "</td>");

	var hk = "<input type='text' name='hk" + lastrow + "' id='hk" + lastrow + "' onkeyup='hanyaAngka(this);getValidasiHK(" + lastrow + ")' value='0' class='ratatengah lebarhk'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + hk + "</td>");

	var bunch = "<input type='text' name='bunch" + lastrow + "' id='bunch" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + bunch + "</td>");

	var jjgu = "<input type='text' name='jjgu" + lastrow + "' id='jjgu" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + jjgu + "</td>");

	var jjgab = "<input type='text' name='jjgab" + lastrow + "' id='jjgab" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + jjgab + "</td>");

	var abw = "<input type='text' name='abw" + lastrow + "' id='abw" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + abw + "</td>");

	var kgbrondolan = "<input type='text' name='kgbrondolan" + lastrow + "' id='kgbrondolan" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + kgbrondolan + "</td>");

	var kgtbs = "<input type='text' name='kgtbs" + lastrow + "' id='kgtbs" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarkgtbs'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + kgtbs + "</td>");

	var output = "<input type='text' name='output" + lastrow + "' id='output" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarabsen'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + output + "</td>");

	var ratepremi = "<input type='text' name='ratekb" + lastrow + "' id='ratekb" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarkgtbs'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + ratepremi + "</td>");

	var allowedcode = "<select name='allowedcode" + lastrow + "' id='allowedcode" + lastrow + "' class='ratakiri lebarallow'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri'>" + allowedcode + "</td>");

	var ratepremi = "<input type='text' name='ratepremi" + lastrow + "' id='ratepremi" + lastrow + "' onkeyup='hanyaAngka(this)' value='0' class='ratatengah lebarrate'>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + ratepremi + "</td>");

	var remark = "<textarea type='text' name='remark" + lastrow + "' id='remark" + lastrow + "' rows='3' cols='10' ></textarea>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + remark + "</td>");
	
	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowUpah(" + lastrow + "," + lastrow + ");' />";
	$('#tru' + lastrow).append("<td>" + tambahbutton + "</td>");

	$('#tru' + lastrow).append("<td>" + kurangbutton + "</td>");

	$('#jumldata').val(lastrow);
	

	showEmployee("fccode" + lastrow);
    showFCBA("lokasi" + lastrow);
	showAbsen("absen" + lastrow);
	showActivity("activity" + lastrow);
	showAllowed("allowedcode" + lastrow);
	getNoRKBFCBA("rkbnofcba" + lastrow);

	new Spry.Widget.ValidationSelect("fccode" + lastrow);
	new Spry.Widget.ValidationSelect("absen" + lastrow);
	new Spry.Widget.ValidationSelect("activity" + lastrow);
	new Spry.Widget.ValidationSelect("tcupah" + lastrow);
	new Spry.Widget.ValidationTextField("bunch" + lastrow);
	new Spry.Widget.ValidationTextField("hk" + lastrow);
	
	

}

//get RKB Lokasi upah
function getNoRKBFCBA(selid) {
	var tglbkm = $('#tglbkm').val();
	var lokasi = $('#lokasi' + selid).val();
	$('#rkbnofcba' + selid).empty();
	$('#rkbnofcba' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getrkb.asp?tglbkm=' + tglbkm + '&fcba=' + lokasi + '&id=' + 1+'&doc='+503,
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
	Spry.Widget.Utils.destroyWidgets('bunch' + i);
	Spry.Widget.Utils.destroyWidgets('jjgu' + i);
	Spry.Widget.Utils.destroyWidgets('jjgab' + i);
	Spry.Widget.Utils.destroyWidgets('abw' + i);
	Spry.Widget.Utils.destroyWidgets('kgbrondolan' + i);
	Spry.Widget.Utils.destroyWidgets('kgtbs' + i);
	Spry.Widget.Utils.destroyWidgets('output' + i);
	Spry.Widget.Utils.destroyWidgets('ratekb' + i);
	Spry.Widget.Utils.destroyWidgets('allowedcode' + i);
	Spry.Widget.Utils.destroyWidgets('ratekb' + i);
	Spry.Widget.Utils.destroyWidgets('ratepremi' + i);
	Spry.Widget.Utils.destroyWidgets('remark' + i);
	
	
	 $('#tru' + i).remove();
	
}

var StartIterationitem=1;
function addrowMutu(trid)
{

	var lstitem = this.StartIterationitem++;
	
	if (lstitem < 2) {
		var row = $("#trheader1").after("<tr id='trb" + lstitem + "'></tr>");
		var kurangbutton = "&nbsp;";
	} else {
		var row = $('#trb' + trid).after("<tr id='trb" + lstitem + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowMutu(" + lstitem + ");' />";
	}

	var tcmutu = "<select name='tcmutu" + lstitem + "' id='tcmutu" + lstitem + "' class='lebartc ratatengah'></select>";
	$('#trb' + lstitem).append("<td class='ratatengah'>" + tcmutu + "</td>");

	var bmateng = "<input type='text' name='bmateng" + lstitem + "' id='bmateng" + lstitem + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratatengah'>" + bmateng + "</td>");

	var blmateng = "<input type='text' name='blmateng" + lstitem + "' id='blmateng" + lstitem + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratatengah'>" + blmateng + "</td>");

	var bkmateng = "<input type='text' name='bkmateng" + lstitem + "' id='bkmateng" + lstitem + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratatengah'>" + bkmateng + "</td>");

	var bmentah = "<input type='text' name='bmentah" + lstitem + "' id='bmentah" + lstitem + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratatengah'>" + bmentah + "</td>");

	var bbusuk = "<input type='text' name='bbusuk" + lstitem + "' id='bbusuk" + lstitem + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratatengah'>" + bbusuk + "</td>");

	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowMutu(" + lstitem + ");' />";
	$('#trb' + lstitem).append("<td>" + tambahbutton + "</td>");

	$('#trb' + lstitem).append("<td>" + kurangbutton + "</td>");

	$('#jumldataitem').val(lstitem);

	showBlokmutu("tcmutu" + lstitem);
	
	
}

//hapus row bahan
function deleterowMutu(i) {
	Spry.Widget.Utils.destroyWidgets('tcmutu' + i);
	Spry.Widget.Utils.destroyWidgets('bmateng' + i);
	Spry.Widget.Utils.destroyWidgets('blmateng' + i);
	Spry.Widget.Utils.destroyWidgets('bkmateng' + i);
	Spry.Widget.Utils.destroyWidgets('bmentah' + i);
	Spry.Widget.Utils.destroyWidgets('bbusuk' + i);
	 $('#trb' + i).remove();
	
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
				$('#timein' + i).html(timein);
				$('#timeout' + i).html(timeout);
				$('#timeinx' + i).val(timeinx);
				$('#timeoutx' + i).val(timeoutx);
			
			
		});
		
	}

}


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
				window.close();
				window.open('bkm_panen_edit.Asp?bkmno=' + bkmno);
			}else{
				getEmployeeview();
			}
			
			
						
		});
		
	} 
}


//untuk input denda panen
var StartIterationdenda=1;
function addrowDenda(trid)
{

	var lsdenda = this.StartIterationdenda++;
	
	if  (lsdenda < 2) {
		var row = $("#trheader3").after("<tr id='trd" + lsdenda + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowDenda(" + lsdenda + ");' />";
	} else {
		var row = $('#trd' + trid).after("<tr id='trd" + lsdenda + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowDenda(" + lsdenda + ");' />";
	}

	var empcode = "<select name='empcode" + lsdenda + "' id='empcode" + lsdenda + "' class='lebaremp ratatengah'></select>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + empcode + "</td>");

	var ktbsa = "<input type='text' name='ktbsa" + lsdenda + "' id='ktbsa" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + ktbsa + "</td>");

	var ktbsba = "<input type='text' name='ktbsba" + lsdenda + "' id='ktbsba" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + ktbsba + "</td>");

	var ktbstp = "<input type='text' name='ktbstp" + lsdenda + "' id='ktbstp" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + ktbstp + "</td>");

	var ktbsbk = "<input type='text' name='ktbsbk" + lsdenda + "' id='ktbsbk" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + ktbsbk + "</td>");

	var ktbso = "<input type='text' name='ktbso" + lsdenda + "' id='ktbso" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + ktbso + "</td>");

	var ktbsab = "<input type='text' name='ktbsab" + lsdenda + "' id='ktbsab" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + ktbsab + "</td>");

	var ktbse = "<input type='text' name='ktbse" + lsdenda + "' id='ktbse" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + ktbse + "</td>");

	var kancakbt = "<input type='text' name='kancakbt" + lsdenda + "' id='kancakbt" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + kancakbt + "</td>");

	var kancakps = "<input type='text' name='kancakps" + lsdenda + "' id='kancakps" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + kancakps + "</td>");

	var kancaksp = "<input type='text' name='kancaksp" + lsdenda + "' id='kancaksp" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + kancaksp + "</td>");

	var kancakbrd = "<input type='text' name='kancakbrd" + lsdenda + "' id='kancakbrd" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + kancakbrd + "</td>");

	var kancakop = "<input type='text' name='kancakop" + lsdenda + "' id='kancakop" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + kancakop + "</td>");

	var kancakbm = "<input type='text' name='kancakbm" + lsdenda + "' id='kancakbm" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + kancakbm + "</td>");

	var kancakbdt = "<input type='text' name='kancakbdt" + lsdenda + "' id='kancakbdt" + lsdenda + "' value='0' onkeyup='hanyaAngka(this)' class='lebartc ratatengah'>";
	$('#trd' + lsdenda).append("<td class='ratatengah'>" + kancakbdt + "</td>");

	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowDenda(" + lsdenda + ");' />";
	$('#trd' + lsdenda).append("<td>" + tambahbutton + "</td>");

	$('#trd' + lsdenda).append("<td>" + kurangbutton + "</td>");

	$('#jumldatadenda').val(lsdenda);

	showEmployee("empcode" + lsdenda);
	
	
}

//hapus row bahan
function deleterowDenda(i) {
	Spry.Widget.Utils.destroyWidgets('empcode' + i);
	Spry.Widget.Utils.destroyWidgets('ktbsa' + i);
	Spry.Widget.Utils.destroyWidgets('ktbsba' + i);
	Spry.Widget.Utils.destroyWidgets('ktbstp' + i);
	Spry.Widget.Utils.destroyWidgets('ktbsbk' + i);
	Spry.Widget.Utils.destroyWidgets('ktbso' + i);
	Spry.Widget.Utils.destroyWidgets('ktbsab' + i);
	Spry.Widget.Utils.destroyWidgets('ktbse' + i);
	Spry.Widget.Utils.destroyWidgets('kancakbt' + i);
	Spry.Widget.Utils.destroyWidgets('kancakps' + i);
	Spry.Widget.Utils.destroyWidgets('kancaksp' + i);
	Spry.Widget.Utils.destroyWidgets('kancakbrd' + i);
	Spry.Widget.Utils.destroyWidgets('kancakop' + i);
	Spry.Widget.Utils.destroyWidgets('kancakbm' + i);
	Spry.Widget.Utils.destroyWidgets('kancakbdt' + i);
	 $('#trd' + i).remove();
	
}


//new Spry.Widget.ValidationTextField("DATECREATED", "none");
//new Spry.Widget.ValidationSelect("clasification");
</script>