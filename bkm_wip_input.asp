<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim businessunit,fcbaname,CompanyName,q,sqltotdata,gangcode,companycode
businessunit= trim(session("fcba"))
gangcode	= trim(request("gangcode"))

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
   companycode=rsCat("fccompanycode")
 
end if



%>

<html>
	<head>
		<title>FBPlants - BKM WIP</title>
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
		width: 90px;
	}
	.lebaractivity{
		width: 190px;
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
		width: 230px;
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
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : BKM WIP</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry" method="post" action="bkm_wip_proses.asp">
	<input type="hidden" name="jumldatawip" id="jumldatawip" value="0" />
	<input type="hidden" name="fccompanycode" id="fccompanycode" value="<%=companycode%>" />
	<table class="data posisikiri" border="0" width="90%">			
		<tr>
			<td align="center" >&nbsp;&nbsp;&nbsp;Tanggal</td>
			<td align="left">
				<input type="text" class="datepicker" maxlength="15" name="tglbkm" readonly id="tglbkm" size="15" value="<%=textDate(date)%>" onchange="get_Allowentry();" title="Tanggal BKM">
			</td>
			<td align="left" >Contractor Name</td>
			<td align="left">
				<select name="contractorcode" id="contractorcode" class="lebarsection" onchange="addrowWIP(1);showFCBA();showSPK();">
				</select>
			</td>
			
		</tr>
		
	</table>
	
	<br>
		
	<table align=center border="0" class='data posisikiri' width='850' id='tableall'>
		<tr>
			<td colspan=>
				<div class="TabbedPanels" id="tp1" align='center'>
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">BKM WIP</li>
					</ul>

					<div class="TabbedPanelsContentGroup">

						<!-- view bkm WIP-->
						<div class="TabbedPanelsContent">
							<!-- ============================================================================================================================= -->
							<div id="datadetail">
								<table align=center border="0" width="850" class='data' id='table4'>
									<tr id="trheader4">
										<th>Lokasi</th>
										<th>SPK No.</th>
										<th>Activity</th>
										<th>Uom</th>
										<th>LineNo</th>
										<th>Rate</th>
										<th>TC</th>
										<th>Qty</th>
										<th>Unit</th>
										<th>Output(Ha)</th>
										<th colspan="2">&nbsp;</th>
									</tr>
									<tr id='trw'>
											<tbody id="bodytablew">
											</tbody>
											
										</tr>
								</table>
							</div>
						</div>


					</div>
				</div>
			</td>
		</tr>
		<tr>
			<td colspan='' align=center><input Type="Submit" id="btnsimpan" name="simpan" value="Simpan"></td>
			
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
		evt.value = evt.value.replace(/,/g, ".");

	}

$(function() {	
	$("#tglbkm").datepicker();
	showContractor();
	
	
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


//get divisi
function showBlok(selid) {
	
	$('#blok').empty();
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


//get contractor
function showContractor() {
	$('#contractorcode').empty();
	var jmlitem = 7;
	$('#contractorcode').select2({
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


function getTCWIP(i)
{
	var activityiwip = $('#activitywip' + i).val();

	$('#tcwip' + i).empty();
	$('#tcwip' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getblok.asp?activity=' + activityiwip + "&tcby=" + 1,
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

function getLine(i)
{
	var spkno = $('#spkno' + i).val();
	var activityiwip = $('#activitywip' + i).val();

	$('#linenowip' + i).empty();
	$('#linenowip' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_show_item.asp?jmlitem=' + 8 + '&activitywip=' + activityiwip + '&agreement=' + spkno,
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
function getRate(i) {
	
	var spkno = $('#spkno' + i).val();
	var activityiwip = $('#activitywip' + i).val();
	var linenowip = $('#linenowip' + i).val();
	
	if (linenowip != "") {
		$('#ratewip' + i).empty();
		$.post('bkm_show_item.asp?jmlitem=' + 9 + '&activitywip=' + activityiwip + '&agreement=' + spkno + '&linenowip=' + linenowip, function(data) {
			var json  = JSON.parse(data);
			var ratewip = json.results[0].rate;
			$("#ratewip" + i).val(ratewip);

		});
		//getTC(activityiwip);
	}	
	
}


function getActivity(i)
{
	var spkno = $('#spkno' + i).val();

	$('#activitywip' + i).empty();
	$('#activitywip' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp?panen=' + 3 + '&spkno=' + spkno,
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

function getTC(i)
{
	var lokasi   = $('#lokasi' + i).val();
	var activityiwip = $('#activitywip' + i).val();
	var tglbkm   = $('#tglbkm').val();
	$('#tcwip' + i).empty();
	$('#tcwip' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_gangcode_wip.asp?gancode=' + 3 +'&lokasi=' + lokasi + '&actcode=' + activityiwip + '&tglbkm=' + tglbkm,
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
	//alert(activityiwip);

}


var StartIterationwip =  1	
function addrowWIP(trid)
{

	var lstwip = this.StartIterationwip++;
	 if (lstwip < 2) {
	 	var row = $("#trheader4").after("<tr id='trw" + lstwip + "'></tr>");
	 	var kurangbutton = "&nbsp;";
	 } else {
		var row = $('#trw' + trid).after("<tr id='trw" + lstwip + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowWIP(" + lstwip + ")' />";
	}

	var lokasi = "<select name='lokasi" + lstwip + "' id='lokasi" + lstwip + "' onchange='getTC(" + lstwip + ")' class='lebarlokasi'></select>";
	$('#trw' + lstwip).append("<td class='ratakiri'>" + lokasi + "</td>");

	var spkno = "<select name='spkno" + lstwip + "' id='spkno" + lstwip + "' class='lebarspkno' onchange='getActivity(" + lstwip + ")'></select>";
	$('#trw' + lstwip).append("<td class='ratakiri'>" + spkno + "</td>");

	var activitywip = "<select name='activitywip" + lstwip + "' id='activitywip" + lstwip + "' onchange='getLine(" + lstwip + ");getuom(" + lstwip + ")' class='lebaractivity'></select>";
	$('#trw' + lstwip).append("<td class='ratakiri'>" + activitywip + "</td>");

	var uom_unit = "<input type='text' name='uom_unit" + lstwip + "' id='uom_unit" + lstwip + "'  value='HA' readonly class='lebartc ratatengah'>";
	$('#trw' + lstwip).append("<td class='ratatengah'>" + uom_unit + "</td>");

	var linenowip = "<select name='linenowip" + lstwip + "' id='linenowip" + lstwip + "' onchange='getRate(" + lstwip + ");getTC(" + lstwip + ")' class='lebartc ratatengah'></select>";
	$('#trw' + lstwip).append("<td class='ratakiri'>" + linenowip + "</td>");

	var ratewip = "<input type='text' name='ratewip" + lstwip + "' id='ratewip" + lstwip + "' onkeyup='hanyaAngka(this)' value='0' readonly class='lebarrate ratakanan'>";
	$('#trw' + lstwip).append("<td class='ratakiri'>" + ratewip + "</td>");

	var tcwip = "<select name='tcwip" + lstwip + "' id='tcwip" + lstwip + "' class='lebartc ratatengah'></select>";
	$('#trw' + lstwip).append("<td class='ratakiri'>" + tcwip + "</td>");

	var qtywip = "<input type='text' name='qtywip" + lstwip + "' id='qtywip" + lstwip + "' onkeyup='hanyaAngka(this)' value='0' class='lebarabsen ratatengah'>";
	$('#trw' + lstwip).append("<td class='ratakiri'>" + qtywip + "</td>");

	var unitwip = "<input type='text' name='unitwip" + lstwip + "' id='unitwip" + lstwip + "' onkeyup='hanyaAngka(this)' value='0' class='lebarabsen ratatengah'>";
	$('#trw' + lstwip).append("<td class='ratakiri'>" + unitwip + "</td>");

	var outputwip = "<input type='text' name='outputwip" + lstwip + "' id='outputwip" + lstwip + "' onkeyup='hanyaAngka(this)' value='0' class='lebarabsen ratatengah'>";
	$('#trw' + lstwip).append("<td class='ratakiri'>" + outputwip + "</td>");

	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowWIP(" + lstwip + ")' />";
	$('#trw' + lstwip).append("<td>" + tambahbutton + "</td>");

	$('#trw' + lstwip).append("<td>" + kurangbutton + "</td>");

	$("#jumldatawip").val(lstwip);

	showSPK("spkno" + lstwip);
	showFCBA("lokasi" + lstwip);
	//showtcWIP("tcwip" + lstwip);
	

	
}

// DELETE ROW WIP
function deleterowWIP(i) {
	Spry.Widget.Utils.destroyWidgets('lokasi' + i);
	Spry.Widget.Utils.destroyWidgets('spkno' + i);
	Spry.Widget.Utils.destroyWidgets('activitywip' + i);
	Spry.Widget.Utils.destroyWidgets('uom_unit' + i);
	Spry.Widget.Utils.destroyWidgets('tcwip' + i);
	Spry.Widget.Utils.destroyWidgets('linenowip' + i);
	Spry.Widget.Utils.destroyWidgets('ratewip' + i);
	Spry.Widget.Utils.destroyWidgets('qtywip' + i);
	Spry.Widget.Utils.destroyWidgets('unitwip' + i);
	Spry.Widget.Utils.destroyWidgets('outputwip' + i);
	$('#trw' + i).remove();
	
}

//ambil data satuan
function getuom(i)
{
	var activity = $('#activitywip' + i).val();
	if (activity != "") {
		$.post("bkm_getactivity.asp?activity=" + activity +"&panen=" + 10, function(data) {
			$('#uom_unit' + i).empty();
			var json  = JSON.parse(data);
			var uom = json.results[0].uom_unit
			$('#uom_unit' + i).val(uom);
						
		});
		
	}
}

//get lokasi upah
function showFCBA(selid) {
	var fcba = 'fcba';
	var companycode = $('#fccompanycode').val();
	$('#lokasi').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getabsen.asp?viewdata=' + fcba + '&fccompanycode=' + companycode,
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

//get WIP
function showSPK(selid) {
	//$('#spkno').empty();

	var jmlitem = 5;
	var tglbkm = $('#tglbkm').val();
	var contractorcode = $('#contractorcode').val();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : "bkm_show_item.asp?jmlitem=" + jmlitem + "&contractorcode=" + contractorcode + "&tglbkm=" + tglbkm,
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
	var tglbkm = $('#tglbkm').val();
	if (tglbkm != "") {
		$.post("bkm_allow_entry.asp?tglbkm=" + tglbkm, function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var closingStatus = json.results[0].msg;
			// if(closingStatus=="FALSE"){
			// 	$('#btnsimpan').show();
			// 	//window.open('bkm_operasional_edit.Asp?bkmno=' + bkmno);
			// }else{
			// 	alert("PERIODE SUDAH DI CLOSE !!!");
			// 	$('#btnsimpan').hide();
			// 	//break;
			// }
			
			
						
		});
		
	} 
}

//new Spry.Widget.ValidationTextField("DATECREATED", "none");
//new Spry.Widget.ValidationSelect("clasification");
</script>