<%validatesession%>
<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include file="DailyValidation.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<html>
<head>
<title>ePlaS - Monthly Work Plan</title><link rel="StyleSheet" href="dtree.css" type="text/css" />
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
		width: 100px;
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
		width: 50px;
	}

	.lebaractbarang{
		width: 180px;
	}
</style>
<%
	dim classification,simpan,sqlhdr,sqldtl,fcba,countrkb
	DIM statuspost,periodethn,periodebln,section,tglinput,ttlamountupah,rkbno
	dim sqlnomor,rsnomor,tglprev,currentnomor,currentnomorfix,chargeunit
	DIM actcode,blok,ttlhk,rate,normamandays,luasdikerjakan,rotasi,totalamounthk,norma_hk
	DIM nQtdFields,nCount,linenokey,tahuntanam
	classification 	= Request("type")
	statuspost 		= Request("statuspost")
	periodethn 		= Request("tahun")
	periodebln 		= Request("bulan")
	section		 	= Request("sectionname")
	tglinput 		= Request("tglinput")
	ttlamountupah 	= Request("amountrkbrpfrm")
	simpan 			= Request("simpan")
	rkbno 			= Request("norkb")
	fcba 			= session("fcba")


	If Len(Trim(tglinput)) = 0 Then	tglinput = textDate(date)

	if simpan="SAVE" then
	
			sqlhdr = "insert into RKBHDR(RKBNO, PERIOD_YEAR, PERIOD_MONTH, SECTIONCODE, CREATEDATE, " & _
			" LASTUPDATE,FCENTRY,FCEDIT,FCIP,FCBA,ROWSTATE,DOCUMENT_CLASSIFICATION,TOTUPAH,CHARGEUNIT) " & _
			" values('" & rkbno & "','" & periodethn & "','" & periodebln & "','" & section & "',TO_DATE('" & tglinput & "', 'mm/dd/yyyy'),SYSDATE," & _ 
			" '" & Session("CodeName") & "', " & _ 
			" '" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "','Planned','503','" & ttlamountupah & "','" & fcba & "')"
			ExecuteAndLog "rkb_panen_input.asp", sqlhdr

			'looping di tabel detail
			nQtdFields 	= request.form("jmlrawat")
			nCount 		= 1
			
			do while cint(nCount) <= cint(nQtdFields) 
				tahuntanam 		= request.form("tahuntanam" & nCount)
				actcode 		= request.form("activityrwt" & nCount)
				blok 			= request.form("tcrwt" & nCount)	
				rate 			= request.form("rate" & nCount)
				bjrbudget		= request.form("bjrbudget" & nCount)
				kgbudget		= request.form("kgbudget" & nCount)
				sb				= request.form("sb" & nCount)
				lb				= request.form("lb" & nCount)
				rateborongan	= request.form("rateborongan" & nCount)
				totalamountborongan	= request.form("amountboronganbck" & nCount)
				norma_hk		= request.form("norma_hk" & nCount)
				rotasi			= request.form("rotasi" & nCount)
				linenokey 		= nCount
				'if qtyrkb > 0 then

				if rateborongan > 0 then 
					rate = 0
					bjrbudget = 0
					sb = 0
					lb = 0
					norma_hk = 0
					rotasi = 0
				end if

					sqldtl = "insert into RKBUPAHDTL(RKBNO, ACTCODE, LOCATIONCODE, RATE, " & _
					" ABW_BUDGET, KG_BUDGET,  " & _
					" RATE_SB,RATE_LB,LINENOKEY, " & _
					" CREATEDATE, LASTUPDATE, FCENTRY, FCEDIT, FCIP,FCBA,ROWSTATE,TAHUNTANAM,RATEBORONGAN,TOTALAMOUNT_BORONGAN,NORMAMANDAYS,ROTASI) " & _
					" values('" & rkbno & "','" & actcode & "','" & blok & "', " & _ 
					" '" & rate & "','" & bjrbudget & "','" & kgbudget & "','" & sb & "','" & lb & "','" & linenokey & "', " & _ 
					" SYSDATE,SYSDATE,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "', " & _ 
					" '" & fcba & "','Planned','" & tahuntanam & "','" & rateborongan & "','" & totalamountborongan & "','" & norma_hk & "','" & rotasi & "')"
					ExecuteAndLog "rkb_panen_input.asp", sqldtl

				'end if

			nCount 	= nCount + 1
			loop
		response.redirect("rkb_panen_new_edit.asp?rkbno="&rkbno)
	End IF
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
	<form name="frmEntry" method="post" action="rkb_panen_new_input.asp">
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
					sqlbulan7 = "select distinct to_number(accountingperiod) bulan,description from accountingperiod order by bulan " 
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
			<td align="left" >Section</td>
			<td align="left" colspan="3">
				<select name="sectionname" id="sectionname" class="lebarsection" onchange="getNorkb();getValidasi();">
						<option value="">--Pilih Section--</option>
				</select>
			</td>
					
		</tr>
		<tr>
			<td align="left">Status</td>
			<td align="left">
				Planned
			</td>
			
		</tr>

	</table>
	
	<br>
	
	<table align=center border="0" class='data posisikiri' width='100%' id='tableall'>
		
		<tr>
			<td>
			   		<div class="TabbedPanels" id="tp1" align='center'>
				
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">RKB Panen</li>
						
					</ul>

					<div class="TabbedPanelsContentGroup">

						<!-- view bkm wip-->
						<div class="TabbedPanelsContent">
							<table align=center border="0" width="900" class='data' id='table2'>
								<tr id="trheader1">
									<th>Tahun Tanam</th>
									<th>Aktivitas</th>
									<th>Blok</th>
									<th>Rate</th>
									<th>BJR Budget</th>
									<th>KG/JJG Budget</th>
									<th>SB(Rp.)</th>
									<th>LB(Rp.)</th>
									<th>Rate Borongan</th>
									<th>Amount Borongan</th>
									<th colspan="2"></th>
								</tr>
								<tr id='trb'>
									<tbody id="bodytablerwt">
									</tbody>
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
	$("#tglinput").datepicker();
	showSection();
	addrowPerawatan();
	

});

function getNorkb()
{
	var sectionname = $("#sectionname").val();
	var tahun 		= $("#tahun").val();
	var bulan 		= $("#bulan").val();
	var fcbatxt 	= $("#fcbatxt").val();
	var classifiction 	= $("#classifiction").val();
	var norkb		= "RKB/" + classifiction + "/" + fcbatxt + "/" + sectionname + "/" + tahun + bulan
	$("#norkb").val(norkb);
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
var StartIterationitem=1;
function addrowPerawatan(trid)
{
	var lstitem = this.StartIterationitem++;
	if (lstitem < 2) {
		var row = $("#trheader1").after("<tr id='trb" + lstitem + "'></tr>");
		var kurangbutton = "&nbsp;";
	} else {
		var row = $('#trb' + trid).after("<tr id='trb" + lstitem + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowBahan(" + lstitem + ");' />";
	}

	var tahuntanam = "<select name='tahuntanam" + lstitem + "' id='tahuntanam" + lstitem + "' class='lebarhaactual' onchange='getNorm(" + lstitem + ");'></select><input type='hidden' name='norma_hk" + lstitem + "' id='norma_hk" + lstitem + "' value='0'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + tahuntanam + "</td>");

	var activityrwt = "<select name='activityrwt" + lstitem + "' id='activityrwt" + lstitem + "' class='lebaractbarang' onchange='getTCrwt(" + lstitem + ");getRate(" + lstitem + ");'></select>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + activityrwt + "</td>");

	var tcrwt = "<select name='tcrwt" + lstitem + "' id='tcrwt" + lstitem + "' class='lebartc ratatengah'></select><input type='hidden' name='rotasi" + lstitem + "' id='rotasi" + lstitem + "' value='0'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + tcrwt + "</td>");

	var rate = "<input type='text' name='rate" + lstitem + "' id='rate" + lstitem + "' value='0' onkeyup='hanyaAngka(this)' style='background-color:#eee;' class='lebarhaactual ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + rate + "</td>");
	
	var bjrbudget = "<input type='text' name='bjrbudget" + lstitem + "' id='bjrbudget" + lstitem + "' value='0' onkeyup='hanyaAngka(this);' class='lebarhaactual ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + bjrbudget + "</td>");
	
	var kgbudget = "<input type='text' name='kgbudget" + lstitem + "' id='kgbudget" + lstitem + "' value='0' onkeyup='hanyaAngka(this);getTotalhk(" + lstitem + ");' class='lebarhaactual ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + kgbudget + "</td>");

	var sb = "<input type='text' name='sb" + lstitem + "' id='sb" + lstitem + "' value='0' onkeyup='hanyaAngka(this);' class='lebarhaactual ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + sb + "</td>");
	
	var lb = "<input type='text' name='lb" + lstitem + "' id='lb" + lstitem + "' value='0' onkeyup='hanyaAngka(this);' class='lebartc ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + lb + "</td>");

	var rateborongan = "<input type='text' name='rateborongan" + lstitem + "' id='rateborongan" + lstitem + "' value='0' onkeyup='hanyaAngka(this);getAmountborongan(" + lstitem + ");' class='lebarhaactual ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + rateborongan + "</td>");

	var amountborongan = "<input type='text' name='amountborongan" + lstitem + "' id='amountborongan" + lstitem + "' style='background-color:#eee;' value='0' onkeyup='hanyaAngka(this)' class='lebarlokasi ratakanan'><input type='hidden' name='amountboronganbck" + lstitem + "' id='amountboronganbck" + lstitem + "' value='0'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + amountborongan + "</td>");
	
	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowPerawatan(" + lstitem + ");' />";
	$('#trb' + lstitem).append("<td>" + tambahbutton + "</td>");

	$('#trb' + lstitem).append("<td>" + kurangbutton + "</td>");

	$('#jmlrawat').val(lstitem);

	showTahuntanam("tahuntanam" + lstitem);
	showActivityrwt("activityrwt" + lstitem);

	
}

//hapus row rkbperawatan
function deleterowBahan(i) {
	Spry.Widget.Utils.destroyWidgets('tahuntanam' + i);
	Spry.Widget.Utils.destroyWidgets('activityrwt' + i);
	Spry.Widget.Utils.destroyWidgets('tcrwt' + i);
	Spry.Widget.Utils.destroyWidgets('rate' + i);
	Spry.Widget.Utils.destroyWidgets('bjrbudget' + i);
	Spry.Widget.Utils.destroyWidgets('kgbudget' + i);
	Spry.Widget.Utils.destroyWidgets('sb' + i);
	Spry.Widget.Utils.destroyWidgets('lb' + i);
	Spry.Widget.Utils.destroyWidgets('rateborongan' + i);
	Spry.Widget.Utils.destroyWidgets('amountborongan' + i);
	Spry.Widget.Utils.destroyWidgets('norma_hk' + i);
	Spry.Widget.Utils.destroyWidgets('rotasi' + i);
	 $('#trb' + i).remove();
	// getjmlRowbhn();
}

//get tahun tanam item
function showTahuntanam(selid) {
	
	$('#tahuntanam').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp?panen=' + 12,
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

//get activity item
function showActivityrwt(selid) {
	
	$('#activityrwt').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp?panen=' + 11,
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

//tampilkan blok
function getTCrwt(i)
{
	var activityrwt = $('#activityrwt' + i).val();
	var sectionname = $('#sectionname').val();
	var tahun = $('#tahun').val();
	var bulan = $('#bulan').val();
	$('#tcrwt' + i).empty();
	$('#tcrwt' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getblok_rkb.asp?activity=' + activityrwt + "&tcby=" + 1 + "&tahun=" + tahun + "&bulan=" + bulan + "&section=" + sectionname,
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



function getRate(i)
{
	var activityrwt = $('#activityrwt' + i).val();
	if (activityrwt != "") {
		$.post("bkm_getactivity.asp?panen=" +  7, function(data) {
			//$('#rate' + i).empty();
			var json  = JSON.parse(data);
			var rate_1 = json.results[0].rate_1
			$('#rate' + i).val(rate_1);
			
						
		});
		
	}
}

function getNorm(i)
{
	var tahuntanam = $('#tahuntanam' + i).val();
	if (tahuntanam != "") {
		$.post("bkm_getactivity.asp?panen=" +  20 + "&tahuntanam=" + tahuntanam, function(data) {
			//$('#rate' + i).empty();
			var json  = JSON.parse(data);
			var norma_hk = json.results[0].norma_hk
			var rotasi = json.results[0].rotasi
			$('#rotasi' + i).val(rotasi);
			$('#norma_hk' + i).val(norma_hk);			
		});
		
	}
}

function getAmountborongan(i)
{ 
	var rateborongan 			= parseFloat($("#rateborongan" + i).val());
	var kgbudget 			= parseFloat($("#kgbudget" + i).val());
	var amountttl 		= (rateborongan*kgbudget);
	amountttl			= Math.round(amountttl);
	$("#amountborongan" + i).val(numformat(amountttl));
	$("#amountboronganbck" + i).val((amountttl));
}

function getValidasi()
{
	var norkb = $("#norkb").val();
	var sectionname = $("#sectionname").val();
	var tahun 		= $("#tahun").val();
	var bulan 		= $("#bulan").val();
	var tiperkb		= 503;
	if (norkb != "") {
		$.post("bkm_getactivity.asp?panen=" +  13 + "&sectionname=" + sectionname + "&tahun=" + tahun + "&bulan=" + bulan + "&tiperkb=" + tiperkb, function(data) {
			var json  = JSON.parse(data);
			var rkbdbno = json.results[0].rkbno
			if(rkbdbno==norkb){
				alert("NO RKB SUDAH ADA,SILAHKAN GANTI DIVISI LAIN ATAU LAKUKAN VIEW DATA..!!");
				$("#norkb").val("");
				$("#sectionname").empty();
			}
		});
		
	}
}

 </script>