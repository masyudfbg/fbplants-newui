<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim pcompany,ptahun
'pcompany = request("Company")
'ptahun   = request("tahun")



%>

<html>
	<head>
		<title>FBPlants - FINMOD</title>
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

<body>
<!--#include file="top.jsp"-->

<br>
<br>
<br>
<br>
<br>


<div id="head_judul" style="float:left;font-size:10px; text-align:left;">
		<strong class="posisikiri">MODULE : BUDGET FINANCIAL MODEL</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry" method="post" action="bud_finmod_proses.asp">
	
	<input type="hidden" name="jumldata" id="jumldata" value="0" />
	<input type="hidden" name="jumldataitem" id="jumldataitem" value="1" />
	<input type="hidden" name="proses" id="proses" value="1" />
	<table class="data posisikiri" border="0" width="90%">			
		<tr>
			<td align="left" >Periode</td>
			<td align="left">
			    <select name="tahunfinmod" id="tahunfinmod" style="width:120px">
				<%
				sqltahun_cls_opr = "select distinct to_number(accountingyear) tahun from accountingyear order by tahun "  
				Set rstahun_cls_opr = Server.CreateObject("ADODB.Recordset")
				rstahun_cls_opr.CursorType = 3
				rstahun_cls_opr.CursorLocation = 3
				response.write sqltahun_cls_opr
				rstahun_cls_opr.Open sqltahun_cls_opr, connstr
				While Not rstahun_cls_opr.EOF 
				    if cint(YEAR(Session("v_Costbook")))=cint(rstahun_cls_opr("tahun")) then %>
					<option value="<%=rstahun_cls_opr("tahun")%>" Selected ><%=rstahun_cls_opr("tahun")%></option>
					<%else%>
					<option value="<%=rstahun_cls_opr("tahun")%>"><%=rstahun_cls_opr("tahun")%></option>
					<%end if%>					
				<%rstahun_cls_opr.movenext
				wend 
				rstahun_cls_opr.close 
				%>	
				</select>
			</td>			
		</tr>
		<tr>
			<td align="left" >Company</td>
				<select name="companyfinmod" id="companyfinmod" onchange="getDetails()">
						<option value=""></option> 
						<% 	q = " SELECT companycode, companyname " & _
							 	" FROM company a INNER JOIN user_company_tab b ON a.companycode=b.fccompanycode " & _
								" WHERE b.fcusercode='" & Session("CodeName") & "'" & _							
							    " order by companyname"
							' response.write q rsCompany,txtCompany
							Set rsCompany = Server.CreateObject("ADODB.Recordset")
							rsCompany.CursorType = 3
							rsCompany.CursorLocation = 3
							rsCompany.Open q, connstr
							if not rsCompany.eof then%>
							<%
							Do While Not rsCompany.eof
							%>
								<option value="<%=rsCompany("companycode")%>" 
									<%if cstr(""&pcompany)=cstr(""&rsCompany("companycode")) then %> selected 
									<%End If%>><%=rsCompany("companycode")%>
								</option> 
							<%rsCompany.movenext
							loop
							End if
							rsCompany.Close
							set rsCompany = Nothing
							q=""
						%>
			</select>
			</td>
		</tr>
	</table>
	
	<br>
		
	<table align=center border="0" class='data posisikiri' cellpadding="0" cellspacing="0" width='1100' id='tableall'>
	<input type='hidden' value='1' name='dataunique' id='dataunique'>
		<tr>
			<td colspan=>
						<table align=center border="0" width="1200" class='data' id='table2'>
								<tr id="trheader2">			
									<th>Type</th>								
									<th>JobGroup.Code</th>
									<th>JobGroup.Name</th>
									<th>Qty</th>
									<th>Priceth>
									<th>Amount</th>
									<th>Remark</th>
								</tr>
								<tr id='tru'>
									<tbody id="bodytable">
									</tbody>
								</tr>

								<tr>
									<td colspan='15' align=center><input Type="Submit" id="SIMPAN" name="simpan" value="SIMPAN"  /> </td>
									
								</tr>
								
							</table>
					
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

	function hanyaAngka(evt) {
		evt.value = evt.value.replace(/[^0-9,.-]/g,"");
	}

$(function() {	
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

//GET EMPLOYEE
function getDetails()
{
	var thnfinmod = $("#tahunfinmod").val();
	var cmpfinmod = $('#companyfinmod').val();
	if (thnfinmod != "")&&(cmpfinmod != "") {
		$.post("bud_finmod_getjobgroup.asp?company=" + cmpfinmod + "&year=" + thnfinmod + "&id=" + 2, function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var jumlview  = json.results.length;
			if(jumlview>0){
				var bkmno = json.results[0].bkmno;
				window.open('bkm_operasional_edit.Asp?bkmno=' + bkmno);
			}else{
				getview();
			}
			
			
						
		});
		
	} 
}

//getEmployee
function getview() {
	
	var gangcode = $("#gangcode").val();
	var tglbkm = $('#tglbkm').val();
	
	if (gangcode != "") {
		$.post("bud_finmod_getjobgroup.asp?id=1", function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var jumlview  = json.results.length;
			var no = 1;
			for (var i = 0; i < jumlview; i++) {
				var fccode       = json.results[i].fccode;
				var fcname       = json.results[i].fcname;
				var fieldstatus = json.results[i].job_forfieldstatus;
				//var hkhi 		= json.results[i].mandayshi;
						
				
				var docketlist = '';
					docketlist += '<tr id="tru'+no+'">';
					docketlist += '<td align="center">' + fieldstatus + '<input type="hidden" class="ratatengah lebarhk" name="t_fieldstatus' + no + '" id="t_fieldstatus' + no + '" value="' + fieldstatus + '"/></td>';
					docketlist += '<td align="center">' + fccode + '<input type="hidden" class="ratatengah lebarhk" name="t_fccode' + no + '" id="t_fccode' + no + '" value="' + fccode + '"/></td>';
					docketlist += '<td align="center">' + fcname + '<input type="hidden" class="ratatengah lebarhk" name="t_fcname' + no + '" id="t_fcname' + no + '" value="' + fcname + '"/></td>';
					docketlist += '<td align="center"><input class="ratatengah lebarhk" type="text" onkeyup="hanyaAngka(this)" name="qty' + no + '" id="qty' + no + '" value="0"/><input class="ratatengah" type="hidden" onkeyup="hanyaAngka(this)" name="ttlhkkhl' + no + '" id="ttlhkkhl' + no + '" value="' + ttlmandays + '"/></td>';
					docketlist += '<td align="center"><input class="ratatengah lebarabsen" type="text" onkeyup="hanyaAngka(this)" name="price' + no + '" id="price' + no + '" value="0"/></td>';
					docketlist += '<td align="center"><input class="ratatengah lebarabsen" type="text" onkeyup="hanyaAngka(this)" name="output' + no + '" id="output' + no + '" value="0"/></td>';
					docketlist += '<td align="center"><input class="ratatengah lebarabsen" type="text" readonly="TRUE" name="amount' + no + '" id="amount' + no + '" value="0"/></td>';		
					docketlist += '<td align="center"><textarea rows="3" cols="10" name="remark' + no + '" id="remark' + no + '"/></textarea></td>';		
					
					docketlist += '</tr>';
					$('#bodytable').append(docketlist);
				
				
				
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
	if(statusemp=="KHL" & parseFloat(ttlhkkhl)>=20 & thnblntglview>20230126){
		alert("KARYAWAN DENGAN STATUS KHL TIDAK BOLEH MELEBIHI DARI 20 HK");
		$('#hk' + i).val(0);
	}else if(hk>1){
		alert("INPUTAN HK TIDAK BOLEH LEBIH DARI 1");
		$('#hk' + i).val(0);
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
function showItem(selid) {
	var tglbkm = $('#tglbkm').val();
	$('#ITEMCODE').empty();
	var jmlitem = 0;
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : "bkm_show_item.asp?jmlitem=" + jmlitem + "&tglbkm=" + tglbkm,
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
	
	$('#activityitem').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp',
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

	var lokasi = "<select name='lokasi" + lastrow + "' id='lokasi" + lastrow + "' class='ratakiri lebarlokasi'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri lebarlokasi'>" + lokasi + "</td>");

	var absen = "<select name='absen" + lastrow + "' id='absen" + lastrow + "' class='ratakiri lebarabsen' onchange='getDisable(" + lastrow + ");'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri lebarabsen'>" + absen + "</td>");

	var activity = "<select name='activity" + lastrow + "' id='activity" + lastrow + "' class='ratakiri lebaractivity' onchange='getTC(" + lastrow + ");getDisable(" + lastrow + ");'></select>";
	$('#tru' + lastrow).append("<td class='ratakiri lebarabsen'>" + activity + "</td>");

	var blok = "<select type='text' name='tcupah" + lastrow + "' id='tcupah" + lastrow + "' value='0' class='ratatengah lebartc'></select>";
	$('#tru' + lastrow).append("<td class='ratatengah'>" + blok + "</td>");

	var hk = "<input type='text' name='hk" + lastrow + "' id='hk" + lastrow + "' onkeyup='hanyaAngka(this);getHKhinew(" + lastrow + ")' value='0' class='ratatengah lebarhk'><input class='ratatengah' type='hidden' onkeyup='hanyaAngka(this)' name='ttlhkkhl" + lastrow + "' id='ttlhkkhl" + lastrow + "' value='0'/>";
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


	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowUpah(" + lastrow + "," + lastrow + ");' />";
	$('#tru' + lastrow).append("<td>" + tambahbutton + "</td>");

	$('#tru' + lastrow).append("<td>" + kurangbutton + "</td>");

	$('#jumldata').val(lastrow);
	

	showEmployee("fccode" + lastrow);
    showFCBA("lokasi" + lastrow);
	showAbsen("absen" + lastrow);
	showActivity("activity" + lastrow);
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

var StartIterationitem=1;
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

	var qtyitem = "<input type='text' name='qtyitem" + lstitem + "' id='qtyitem" + lstitem + "' value='0' onkeyup='hanyaAngka(this)' class='lebarabsen ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + qtyitem + "</td>");

	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowBahan(" + lstitem + ");' />";
	$('#trb' + lstitem).append("<td>" + tambahbutton + "</td>");

	$('#trb' + lstitem).append("<td>" + kurangbutton + "</td>");

	$('#jumldataitem').val(lstitem);

	showItem("ITEMCODE" + lstitem);
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