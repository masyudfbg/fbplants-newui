<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<html>
<head>
<title>ePlaS - RKB Traksi</title><link rel="StyleSheet" href="dtree.css" type="text/css" />
<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
<link rel="StyleSheet" href="dtree.css" type="text/css" />
<!-- HADIR/ font ICON -->
<link rel="stylesheet" href="css/font/font-awesome.css" />
<!-- END:HADIR -->
<script type="text/javascript" src="dtree.js"></script>
<link href="plugins/tabcontent/tabcontent.css" rel="stylesheet" type="text/css" />
<link href="plugins/chosen/chosen.min.css" rel="stylesheet" />

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
	

	.lebarsection{
		width: 160px;
	}
	.lebarnormahk{
		width: 80px;
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

	.lebarlokasi{
		width: 110px;
	}

	.lebarhaactual{
		width: 70px;
	}

	.lebartc{
		width: 80px;
	}

	.lebaractbarang{
		width: 200px;
	}

	.lebaractbarang1{
		width: 200px;
	}

	.lebaritembarang{
		width: 200px;
	}
	
	.lebaruom{
		width: 50px;
	}
	
</style>
<%
	dim classification,simpan,sqlhdr,sqldtl,fcba,countrkb
	DIM statuspost,periodethn,periodebln,section,tglinput,ttlamountupah,rkbno
	dim sqlnomor,rsnomor,tglprev,currentnomor,currentnomorfix,chargeunit
	DIM actcode,blok,ttlhk,rate,normamandays,luasdikerjakan,rotasi,totalamounthk
	DIM nQtdFields,nCount,linenokey,kendaraaan,kmhm

	classification 	= Request("type")
	statuspost 		= Request("statuspost")
	periodethn 		= Request("tahun")
	periodebln 		= Request("bulan")
	kendaraan		= Request("kendaraan")
	tglinput 		= Request("tglinput")
	kmhm 			= Request("kmhm")
	simpan 			= Request("simpan")
	rkbno 			= Request("norkb")
	fcba 			= session("fcba")

	if simpan="SAVE" then
	

			sqlhdr = "insert into RKBHDR(RKBNO, PERIOD_YEAR, PERIOD_MONTH, SECTIONCODE, CREATEDATE, " & _
			" LASTUPDATE,FCENTRY,FCEDIT,FCIP,FCBA,ROWSTATE,DOCUMENT_CLASSIFICATION,TOTUPAH,CHARGEUNIT,KMHM) " & _
			" values('" & rkbno & "','" & periodethn & "','" & periodebln & "','" & kendaraan & "',SYSDATE,SYSDATE," & _ 
			" '" & Session("CodeName") & "', " & _ 
			" '" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "','Planned','502','" & ttlamountupah & "','" & fcba & "','" & kmhm & "')"
			ExecuteAndLog "rkb_traksi_input.asp", sqlhdr

			'looping di tabel detail
			nQtdFields 	= request.form("jmlrawat")
			nCount 		= 1
			
			do while cint(nCount) <= cint(nQtdFields) 
				actcode 		= request.form("activityrwt" & nCount)
				uom 			= request.form("uom" & nCount)
				jarak 			= request.form("jarak" & nCount)
				qty 			= request.form("qty" & nCount)
				rate 			= request.form("rate" & nCount)
				amount			= request.form("amountbck" & nCount)
				
				
				if jarak= "" then
					jarak=0
				end if
				if qty= "" then
					qty=0
				end if
				if rate= "" then
					rate=0
				end if
				if amount= "" then
					amount=0
				end if

				
				linenokey 		= nCount
				'if qtyrkb > 0 then

					sqldtl = "insert into RKBVRADTL(RKBNO, ACTCODE, JARAK, QTY, RATE, " & _
					" AMOUNT, FCENTRY, FCEDIT,FCBA,LASTUPDATE,LINENO) " & _
					" values('" & rkbno & "','" & actcode & "','" & jarak & "','" & qty & "', " & _ 
					" '" & rate & "','" & amount & "', " & _ 
					" '" & Session("CodeName") & "','" & Session("CodeName") & "','" & fcba & "',sysdate, " & _ 
					" '" & nCount & "')"
					ExecuteAndLog "rkb_traksi_input.asp", sqldtl

				'end if

			nCount 	= nCount + 1
			loop
		response.redirect("rkb_traksi_edit.asp?rkbno="&rkbno)

	end if 

		
%>

<body>
<!--#include file="top.jsp"-->
<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
 <!--#include file="menu.asp"-->
</div>
<div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 	<div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
 <div class="box_both" style="overflow:scroll;">  
	<form name="frmEntry" method="post" action="rkb_traksi_input.asp">
	<input type="hidden" name="fcbatxt" readonly id="fcbatxt" value="<%=fcba%>">
	<input type="hidden" name="jmlrawat" id="jmlrawat" value="1" />
	<table align=center border="0" class='data posisikiri' width='100%' id='tableall'>			
		<tr>
			<td  align="left">
				<font face="Verdana" size="1" align="LEFT">No. RKB</font>
			</td>
			<td  align="left">
				<input type="text" name="norkb" readonly id="norkb" size="30" value="" title="No RKB" style='background-color:#eee;'>
			</td>
			<td align="left" >Periode</td>
			<td align="left">
			    <select name="tahun" id="tahun" style="width:100px" onchange="getNorkb()">
				<%
					dim sqltahun7,rstahun7
					sqltahun7 = "select distinct to_number(accountingyear) tahun from accountingyear order by tahun desc"  
					Set rstahun7 = Server.CreateObject("ADODB.Recordset")
					rstahun7.CursorType = 3
					rstahun7.CursorLocation = 3
					rstahun7.Open sqltahun7, connstr
					While Not rstahun7.EOF 
						if cint(year(Date))=cint(rstahun7("tahun")) then %>
						<option value="<%=rstahun7("tahun")%>" Selected ><%=rstahun7("tahun")%></option>
						<%else%>
						<option value="<%=rstahun7("tahun")%>"><%=rstahun7("tahun")%></option>
						<%end if%>					
					<%rstahun7.movenext
					wend 
					rstahun7.close 
					

				%>	
				</select>
				<select name="bulan" id="bulan" style="width:120px" onchange="getNorkb()">
				<%
					dim sqlbulan7,rsbulan7
					sqlbulan7 = "select distinct to_number(accountingperiod) bulan,description from accountingperiod where to_number(accountingperiod) <=12 order by bulan " 
					Set rsbulan7 = Server.CreateObject("ADODB.Recordset")
					rsbulan7.CursorType = 3
					rsbulan7.CursorLocation = 3
					rsbulan7.Open sqlbulan7, connstr
					While Not rsbulan7.EOF 
						if cint(month(Date))=cint(rsbulan7("bulan")) then %>
						<option value="<%=rsbulan7("bulan")%>" Selected><%=rsbulan7("description")%></option>
						<%else%>
						<option value="<%=rsbulan7("bulan")%>"><%=rsbulan7("description")%></option>
						<%end if%>
					<%rsbulan7.movenext
					wend 
					rsbulan7.close 
					

				%>	
				</select>
			</td>
		</tr>
		<tr>
			<td align="left" >Classification</td>
			<td align="left">
				<input type="text" name="classifiction" readonly id="classifiction" size="30" value="<%=classification%>" title="Classification">
				
			</td>
			<td align="left" >Kendaraan</td>
			<td align="left" colspan="3">
				<select name="kendaraan" id="kendaraan" class="lebarsection" onchange="getNorkb();getValidasi()" >
						<option value="">--Pilih Kendaraan--</option>
				</select>
			</td>
		</tr>
		<tr>
			<td align="left">Status</td>
			<td align="left">
				Planned
			</td>
			<td align="left">KM/ HM</td>
			<td align="left">
				<input type="text" name="kmhm" id="kmhm" size="10" value="0" title="KM / HM" class="ratakanan" onkeyup=hanyaAngka(this);>
			</td>
		</tr>

	</table>
	
	<br>
		
	<table align=center border="0" class='data posisikiri' width='100%' id='tableall'>
		
		<tr>
			<td>
			   	<div class="TabbedPanels" id="tp1" align='center'>
				
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">RKB VRA</li>
						
					</ul>	
						<div class="TabbedPanelsContentGroup">

							<!-- view bkm wip-->
							<div class="TabbedPanelsContent">
								<table align=center border="0" width="900" class='data' id='table2'>
									<tr id="trheader1">
										<th>Aktivitas</th>
										<th>UOM</th>
										<th>Norm. Jarak</th>
										<th>Norm. QTY</th>
										<th>Rate</th>
										<th>Amount</th>
										<th colspan="2"></th>
									</tr>
									<tr id='trb'>
										<tbody id="bodytablerwt">
										</tbody>
									</tr>

									<tr class="ratakanan">
											<td colspan="5" class="tebal">Total</td>
											<td class="tebal">
												<input class="ratakanan lebarlokasi" type="text" readonly  style="background-color:#eee;" name="grandtotalfrm" id="grandtotalfrm" value="0"/>
												<input class="ratakanan lebarlokasi" type="hidden" readonly  style="background-color:#eee;" name="grandtotal" id="grandtotal" value="0"/>
												
											</td>
									</tr>
									<tr>
										<td  colspan="10" class="ratatengah"><input type=Submit name="simpan" value="SAVE"></td>
									</tr>
								</table>
							</div>

						</div>
				</div>
			</td>
		</tr>
		
	</table>

</form>
 </div>
 </div>
 </body>
 </head>
 <link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
<script src="/js/select2/select2.min.js" type="text/javascript"></script>
<script type="text/javascript">

new Spry.Widget.TabbedPanels("tp1", {
		defaultTab: 0
	});

function hanyaAngka(evt) {
	evt.value = evt.value.replace(/[^0-9,.-]/g,"");
}

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

$(function() {	
	
	showKendaraan();
	addrowVRA();

	
	$(this).on("keyup", "input[name^='rate']", function(){
    	let subtotal =  _hitungsubtotal();
		//alert('tes');
  	});

	
});

function getValidasi()
{
	var norkb = $("#norkb").val();
	var kendaraaan = $("#kendaraan").val();
	var tahun 		= $("#tahun").val();
	var bulan 		= $("#bulan").val();
	var tiperkb		= 502;
	if (norkb != "") {
		$.post("bkm_getactivity.asp?panen=" +  13 + "&sectionname=" + kendaraaan + "&tahun=" + tahun + "&bulan=" + bulan + "&tiperkb=" + tiperkb, function(data) {
			var json  = JSON.parse(data);
			var rkbdbno = json.results[0].rkbno
			if(rkbdbno==norkb){
				alert("NO RKB SUDAH ADA,SILAHKAN GANTI DIVISI LAIN ATAU LAKUKAN VIEW DATA..!!");
				$("#norkb").val("");
				$("#kendaraan").empty();
			}
		});
		
	}
}

function getNorkb()
{
	 var kendaraaan = $("#kendaraan").val();
	var tahun 		= $("#tahun").val();
	var bulan 		= $("#bulan").val();
	var fcbatxt 	= $("#fcbatxt").val();
	var classifiction 	= $("#classifiction").val();
	var norkb		= "RKB/" + classifiction + "/" + fcbatxt + "/" + kendaraaan + "/" + tahun + bulan
	$("#norkb").val(norkb);
	//alert(kendaraaan);
}


 //get section
function showKendaraan() {
	
	$('#kendaraan').empty();
	$('#kendaraan').select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getsection.asp?section=' + 10,
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

//show fcba
function getFCBA()
{
	var company = $("#companyar").val();
	showFCBA(company);
}

function showFCBA(company) {
	$('#fcba').empty();
	$('#fcba').select2({
		ajax: {
			delay: 1000,
			url : 'getListfcbaallajax.asp?fccompanycode=' + company,
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

//rkb perawaran add
var StartIteration=1;
function addrowVRA(trid)
{
	var lstitem = this.StartIteration++;
	if (lstitem < 2) {
		var row = $("#trheader1").after("<tr id='trb" + lstitem + "'></tr>");
		var kurangbutton = "&nbsp;";
	} else {
		var row = $('#trb' + trid).after("<tr id='trb" + lstitem + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowBahan(" + lstitem + ");' />";
	}

	var activityrwt = "<select name='activityrwt" + lstitem + "' id='activityrwt" + lstitem + "' class='lebaractbarang' onchange='getUOM(" + lstitem + ");getNormajarak(" + lstitem + ");getQty(" + lstitem + ");'></select>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + activityrwt + "</td>");

	var uom = "<input type='text' name='uom" + lstitem + "' id='uom" + lstitem + "' style='background-color:#eee;' value='' readonly class='lebarhaactual ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratatengah'>" + uom + "</td>");

	var jarak = "<select name='jarak" + lstitem + "' id='jarak" + lstitem + "' class='lebartc ratatengah' ></select>";
	$('#trb' + lstitem).append("<td class='ratatengah'>" + jarak + "</td>");

	var qty = "<select name='qty" + lstitem + "' id='qty" + lstitem + "' class='lebartc ratatengah' ></select>";
	$('#trb' + lstitem).append("<td class='ratatengah'>" + qty + "</td>");

	var rate = "<input type='text' name='rate" + lstitem + "' id='rate" + lstitem + "' value='0' onkeyup='hanyaAngka(this);getAmount(" + lstitem + ");' class='lebarhaactual ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakanan'>" + rate + "</td>");

	var amount = "<input type='text' name='amount" + lstitem + "' id='amount" + lstitem + "' style='background-color:#eee;' value='0' onkeyup='hanyaAngka(this)'  class='amount lebarlokasi ratakanan'><input type='hidden' name='amountbck" + lstitem + "' id='amountbck" + lstitem + "' value='0'>";
	$('#trb' + lstitem).append("<td class='ratakanan'>" + amount + "</td>");

	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowVRA(" + lstitem + ");' />";
	$('#trb' + lstitem).append("<td>" + tambahbutton + "</td>");

	$('#trb' + lstitem).append("<td>" + kurangbutton + "</td>");

	$('#jmlrawat').val(lstitem);

	showActivityrwt("activityrwt" + lstitem);

	
}

//hapus row rkbperawatan
function deleterowBahan(i) {
	Spry.Widget.Utils.destroyWidgets('activityrwt' + i);
	Spry.Widget.Utils.destroyWidgets('uom' + i);
	Spry.Widget.Utils.destroyWidgets('jarak' + i);
	Spry.Widget.Utils.destroyWidgets('rate' + i);
	Spry.Widget.Utils.destroyWidgets('qty' + i);
	Spry.Widget.Utils.destroyWidgets('amount' + i);
	 $('#trb' + i).remove();
}

//get activity vra
function showActivityrwt(selid) {
	
	$('#activityrwt').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp?panen=' + 16,
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


function getUOM(i)
{
	var activityrwt = $('#activityrwt' + i).val();
	if (activityrwt != "") {
		$.post("bkm_getactivity.asp?activity=" + activityrwt + "&panen=" +  17, function(data) {
			$('#uom' + i).empty();
			var json  = JSON.parse(data);
			var uom = json.results[0].uom
			$('#uom' + i).val(uom);
						
		});
		
	}
}


//tampilkan norma  jarak
function getNormajarak(i)
{
	var activityrwt = $('#activityrwt' + i).val();
	$('#jarak' + i).empty();
	$('#jarak' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp?activity=' + activityrwt + "&panen=" + 18,
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

//tampilkan norma qty 
function getQty(i)
{
	var activityrwt = $('#activityrwt' + i).val();
	$('#qty' + i).empty();
	$('#qty' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp?activity=' + activityrwt + "&panen=" + 19,
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

function getAmount(i)
{ 
	var rate 			= parseFloat($("#rate" + i).val());
	var qty 			= parseFloat($("#qty" + i).val());
	var amountttl 		= (rate*qty);
	amountttl			= Math.round(amountttl);
	$("#amount" + i).val(numformat(amountttl));
	$("#amountbck" + i).val((amountttl));
	
}


var _hitungsubtotal = () => {
    let tsubtotal = 0;
    $('.amount').each(function(index,element) {
      tsubtotal += Number($("input[name^='amountbck']").eq(index).val());     
    });  
  
    tsubtotal = tsubtotal.toString().replace('.',',');            
  
    if(tsubtotal==0) tsubtotal='0,00';

	let grandtotal = numformat(tsubtotal);
    $('#grandtotal').val(tsubtotal).attr('placeholder',tsubtotal);      
    $('#grandtotalfrm').val(grandtotal).attr('placeholder',grandtotal);      
    return;
  }


 </script>