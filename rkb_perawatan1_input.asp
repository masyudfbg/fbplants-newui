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
		width: 180px;
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
	DIM nQtdFields,nCount,linenokey
	DIM rateborongan,totalamountborongan,grandttlborongan,totalsemua,ttlbahan
	classification 	= Request("type")
	statuspost 		= Request("statuspost")
	periodethn 		= Request("tahun")
	periodebln 		= Request("bulan")
	section		 	= Request("sectionname")
	tglinput 		= Request("tglinput")
	ttlamountupah 	= Request("amountrkbrp")
	grandttlborongan = Request("amountborrkbrp")
	ttlbahan 		= Request("amountbahan")
	simpan 			= Request("simpan")
	rkbno 			= Request("norkb")
	fcba 			= session("fcba")

	totalsemua		= cdbl(ttlamountupah) + cdbl(grandttlborongan)

	if simpan="SAVE" then
	
		' buat nomor transaksi'
		sqlnomor = "SELECT COUNT(RKBNO) RKBNO " & _
		" FROM RKBHDR " & _
		" WHERE RKBNO='" & rkbno & "' " & _
		" AND FCBA='" & fcba & "'"
		Set rsnomor = Server.CreateObject("ADODB.Recordset")
		rsnomor.CursorType = 3
		rsnomor.CursorLocation = 3
		rsnomor.Open sqlnomor, connstr
		if not rsnomor.eof then
				countrkb = rsnomor("rkbno")
		end if
		rsnomor.close 
		set rsnomor = nothing

		'if countrkb> 0 then
			'Response.Write "NORKB SUDAH ADA TERSIMPAN"
		'else

			sqlhdr = "insert into RKBHDR(RKBNO, PERIOD_YEAR, PERIOD_MONTH, SECTIONCODE, CREATEDATE, " & _
			" LASTUPDATE,FCENTRY,FCEDIT,FCIP,FCBA,ROWSTATE,DOCUMENT_CLASSIFICATION,TOTUPAH,CHARGEUNIT) " & _
			" values('" & rkbno & "','" & periodethn & "','" & periodebln & "','" & section & "',SYSDATE,SYSDATE," & _ 
			" '" & Session("CodeName") & "', " & _ 
			" '" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "','Planned','501','" & totalsemua & "','" & fcba & "')"

			ExecuteAndLog "rkb_perawatan_input.asp", sqlhdr

			'looping di tabel detail
			nQtdFields 	= request.form("jmlrawat")
			nCount 		= 1
			
			do while cint(nCount) <= cint(nQtdFields) 
				actcode 		= request.form("activityrwt" & nCount)
				blok 			= request.form("tcrwt" & nCount)
				haactual 		= request.form("haactual" & nCount)
				ttlhk 			= request.form("ttlhk" & nCount)
				rate 			= request.form("rate" & nCount)
				normamandays	= request.form("mandayform" & nCount)
				luasdikerjakan	= request.form("luasdikerjakan" & nCount)
				rotasi			= request.form("rotasi" & nCount)
				totalamounthk	= request.form("amountbck" & nCount)
				rateborongan	= request.form("rateborongan" & nCount)
				totalamountborongan	= request.form("amountboronganbck" & nCount)
				if ttlhk= "" then
					ttlhk=0
				end if
				if haactual= "" then
					haactual=0
				end if
				if rate= "" then
					rate=0
				end if
				if normamandays= "" then
					normamandays=0
				end if
				if luasdikerjakan= "" then
					luasdikerjakan=0
				end if
				if rotasi= "" then
					rotasi=0
				end if
				if totalamounthk= "" then
					totalamounthk=0
				end if

				if rateborongan > 0 then 
					ttlhk = 0
				end if
				
				linenokey 		= nCount
				'if qtyrkb > 0 then

					sqldtl = "insert into RKBUPAHDTL(RKBNO, ACTCODE, LOCATIONCODE, MANDAYS, RATE, " & _
					" NORMAMANDAYS, VOLUME,  " & _
					" ROTASI,TOTALAMOUNT_HK,LINENOKEY, " & _
					" CREATEDATE, LASTUPDATE, FCENTRY, FCEDIT, FCIP,FCBA,HAACTUAL,ROWSTATE,RATEBORONGAN,TOTALAMOUNT_BORONGAN) " & _
					" values('" & rkbno & "','" & actcode & "','" & blok & "','" & ttlhk & "', " & _ 
					" '" & rate & "','" & normamandays & "','" & luasdikerjakan & "','" & rotasi & "','" & totalamounthk & "','" & linenokey & "', " & _ 
					" TO_DATE('" & tglinput & "', 'mm/dd/yyyy'),SYSDATE,'" & Session("CodeName") & "','" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "', " & _ 
					" '" & fcba & "','" & haactual & "','Planned','" & rateborongan & "','" & totalamountborongan & "')"
					ExecuteAndLog "rkb_perawatan_input.asp", sqldtl

				'end if

			nCount 	= nCount + 1
			loop

				
		response.redirect("rkb_perawatan1_edit.asp?rkbno="&rkbno)

	end if 

	' simpan bahan
	if simpan = "SAVE BAHAN" then
		dim sqlbahan,nQtdFieldsbahan,nCountbahan,linenokeybahan

		' buat nomor transaksi'
		sqlnomor = "SELECT COUNT(RKBNO) RKBNO " & _
		" FROM RKBHDR " & _
		" WHERE RKBNO='" & rkbno & "' " & _
		" AND FCBA='" & fcba & "'"
		Set rsnomor = Server.CreateObject("ADODB.Recordset")
		rsnomor.CursorType = 3
		rsnomor.CursorLocation = 3
		rsnomor.Open sqlnomor, connstr
		if not rsnomor.eof then
				countrkb = rsnomor("rkbno")
		end if
		rsnomor.close 
		set rsnomor = nothing

		if countrkb="0" then
			sqlhdr = "insert into RKBHDR(RKBNO, PERIOD_YEAR, PERIOD_MONTH, SECTIONCODE, CREATEDATE, " & _
			" LASTUPDATE,FCENTRY,FCEDIT,FCIP,FCBA,ROWSTATE,DOCUMENT_CLASSIFICATION,TOTUPAH,CHARGEUNIT) " & _
			" values('" & rkbno & "','" & periodethn & "','" & periodebln & "','" & section & "',SYSDATE,SYSDATE," & _ 
			" '" & Session("CodeName") & "', " & _ 
			" '" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "','Planned','501','" & totalsemua & "','" & fcba & "')"
			ExecuteAndLog "rkb_perawatan_input.asp", sqlhdr
		end if

			nQtdFieldsbahan 	= request.form("jumldataitem")
			nCountbahan 		= 1
			linenokeybahan 		= 1	

			do while cint(nCountbahan) <= cint(nQtdFieldsbahan)

				dim itembarang 
				itembarang = request.form("itemcode" & nCountbahan)
				if itembarang<>"" then 

					sqlbahan = " insert into RKBBAHANDTL (RKBNO,ACTCODE,LOCATIONCODE,STOCKCODE, " & _ 
					" UOM,QTY,CREATEDATE,LASTUPDATE,FCENTRY, " & _
					" FCEDIT,FCIP,FCBA,LINENOKEY,RATESTOCK,TOTALAMOUNT,ACTIVEMATERIALNORM) " & _
					" values ('" & rkbno & "','" & request.form("activityitem" & nCountbahan) & "','" & request.form("tcitem" & nCountbahan) & "', " & _ 
					" '" & request.form("itemcode" & nCountbahan) & "','" & request.form("uom" & nCountbahan) & "', " & _ 
					" '" & request.form("qtyitem" & nCountbahan) & "',SYSDATE,SYSDATE,'" & Session("CodeName") & "','" & Session("CodeName") & "', " & _ 
					" '" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcba & "','" & linenokeybahan & "','" & request.form("rateitem" & nCountbahan) & "', " & _ 
					" '" & request.form("amountitemx" & nCountbahan) & "','" & request.form("normamm" & nCountbahan) & "')"
					ExecuteAndLog "rkb_perawatan_input.asp", sqlbahan

				end if

 			nCountbahan 	= nCountbahan + 1
			linenokeybahan 	= linenokeybahan + 1

			loop
			sqldtl = "update rkbhdr set totbahan='" & ttlbahan & "' where rkbno='" & rkbno & "'"
			connStr.execute(sqldtl)
			response.redirect("rkb_perawatan1_edit.asp?rkbno="&rkbno)
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
	<form name="frmEntry" method="post" action="rkb_perawatan1_input.asp">
	<input type="hidden" name="fcbatxt" readonly id="fcbatxt" value="<%=fcba%>">
	<input type="hidden" name="jmlrawat" id="jmlrawat" value="1" />
	<input type="hidden" name="jumldataitem" id="jumldataitem" value="1" />
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
				Status
			</td>
			
		</tr>

	</table>
	
	<br>
		<input type="hidden" name="jmlrkb" id="jmlrkb" value="0" />
	<table align=center border="0" class='data posisikiri' width='100%' id='tableall'>
		
		<tr>
			<td>
			   	<div class="TabbedPanels" id="tp1" align='center'>
				
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">RKB Perawatan</li>
						<!--li class="TabbedPanelsTab">RKB Bahan</li-->
					</ul>	
						<div class="TabbedPanelsContentGroup">

						
							<div class="TabbedPanelsContent">
								<table align=center border="0" width="900" class='data' id='table2'>
									<tr id="trheader1">
										<th>Aktivitas</th>
										<th>Blok</th>
										<th>Ha Actual</th>
										<th>Rate</th>
										<th>Mandays Norm</th>
										<th>Luas Dikerjakan</th>
										<th>Total HK</th>
										<th>Rotasi</th>
										<th>Amount</th>
										<th>Rate Borongan</th>
										<th>Amount Borongan</th>
										<th colspan="2"></th>
									</tr>
									<tr id='trb'>
										<tbody id="bodytablerwt">
										</tbody>
									</tr>

									<tr class="ratakanan">
											<td colspan="5" class="tebal">Total</td>
											<td class="tebal">
												<input class="ratatengah lebarhaactual" type="text" readonly  style="background-color:#eee;" name="amountluasdikerjakanfrm" id="amountluasdikerjakanfrm" value="0"/>
												<input class="ratatengah lebarhaactual" type="hidden" readonly  style="background-color:#eee;" name="amountluasdikerjakan" id="amountluasdikerjakan" value="0"/>
											
											</td>
											<td class="tebal">
												<input class="ratakanan lebarhaactual" type="text" readonly  style="background-color:#eee;" name="amountttlhkfrm" id="amountttlhkfrm" value="0"/>
												<input class="ratakanan lebarhaactual" type="hidden" readonly  style="background-color:#eee;" name="amountttlhk" id="amountttlhk" value="0"/>
											</td>
											<td class="tebal">
												<input class="ratakanan lebarhaactual" type="text" readonly  style="background-color:#eee;" name="amountrotasifrm" id="amountrotasifrm" value="0"/>
												<input class="ratakanan lebarhaactual" type="hidden" readonly  style="background-color:#eee;" name="amountrotasi" id="amountrotasi" value="0"/>
											</td>
											<td class="tebal">
												<input class="ratakanan lebarlokasi" type="text" readonly  style="background-color:#eee;" name="amountrkbrpfrm" id="amountrkbrpfrm" value="0"/>
												<input class="ratakanan lebarlokasi" type="hidden" readonly  style="background-color:#eee;" name="amountrkbrp" id="amountrkbrp" value="0"/>
											</td>
											<td class="tebal">
												
											</td>
											<td class="tebal">
												<input class="ratakanan lebarlokasi" type="text" readonly  style="background-color:#eee;" name="amountborrkbrpfrm" id="amountborrkbrpfrm" value="0"/>
												<input class="ratakanan lebarlokasi" type="hidden" readonly  style="background-color:#eee;" name="amountborrkbrp" id="amountborrkbrp" value="0"/>
											</td>
									</tr>
									<tr>
										<td  colspan="10" class="ratatengah"><input type=Submit name="simpan" value="SAVE"></td>
									</tr>
								</table>
							</div>

							<div class="TabbedPanelsContent">
									<!-- ============================================================================================================================= -->
									<!--div id="datadetail">
										<table align=center border="0"  width="850" class='data' id='table1'>
											<tr id="trheader2">
												<th>Activity</th>
												<th>TC</th>
												<th>Item</th>
												<th>Uom</th>
												<th>Norm</th>
												<th>Qty</th>
												<th>Rate</th>
												<th>Amount</th>
												<th colspan="2">&nbsp;</th>
											</tr>

											<tr id='trb1'>
												<tbody id="bodytablebrg">
												</tbody>
												
											</tr>

											<tr class="ratakanan">
												<td colspan="5" class="tebal">Total Bahan</td>
												<td class="tebal">
													<input class="ratatengah lebarhaactual" type="text" readonly  style="background-color:#eee;" name="ttlqtyfrm" id="ttlqtyfrm" value="0"/>
													<input class="ratatengah lebarhaactual" type="hidden" readonly  style="background-color:#eee;" name="ttlqty" id="ttlqty" value="0"/>
												
												</td>
												<td class="tebal">
													
												</td>
												<td class="tebal">
													<input class="ratakanan lebarhaactual" type="text" readonly  style="background-color:#eee;" name="amountbahanfrm" id="amountbahanfrm" value="0"/>
													<input class="ratakanan lebarhaactual" type="hidden" readonly  style="background-color:#eee;" name="amountbahan" id="amountbahan" value="0"/>
												</td>
												
											</tr>

											<tr>
												<td colspan='8' align=center><input Type="Submit" id="SIMPAN_BAHAN" name="simpan" value="SAVE BAHAN"  /></td>
												
											</tr>
											
										</table>
									</div-->
									<!-- ============================================================================================================================= -->
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
	$("#tglinput").datepicker();
	showSection();
	addrowPerawatan();
	addrowBahan(1);

	$(this).on("keyup", "input[name^='luasdikerjakan']", function(){
    	 _hitungLuasdikerjakan();
		 _hitungTtlhk();
		 _hitungTtlrotasi();
		
  	});

	$(this).on("keyup", "input[name^='rotasi']", function(){
		 _hitungTtlrotasi();
		_hitungTtlAmount();
  	});
	
	$(this).on("keyup", "input[name^='rateborongan']", function(){
		 _hitungAmountBorongan();
  	});

	$(this).on("keyup", "input[name^='qtyitem']", function(){
		 _hitungQtyBahan();
		// _hitungAmountBahan();
  	});

	$(this).on("keyup", "input[name^='rateitem']", function(){
		// _hitungQtyBahan();
		 _hitungAmountBahan();
  	});
	
	
});

function getValidasi()
{
	var norkb = $("#norkb").val();
	var sectionname = $("#sectionname").val();
	var tahun 		= $("#tahun").val();
	var bulan 		= $("#bulan").val();
	var tiperkb		= 501;
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
var StartIteration=1;
function addrowPerawatan(trid)
{
	var lstitem = this.StartIteration++;
	if (lstitem < 2) {
		var row = $("#trheader1").after("<tr id='trb" + lstitem + "'></tr>");
		var kurangbutton = "&nbsp;";
	} else {
		var row = $('#trb' + trid).after("<tr id='trb" + lstitem + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowBahan(" + lstitem + ");' />";
	}

	var activityrwt = "<select name='activityrwt" + lstitem + "' id='activityrwt" + lstitem + "' class='lebaractbarang' onchange='getTCrwt(" + lstitem + ");getRate(" + lstitem + ");getNormamandays(" + lstitem + ")'></select>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + activityrwt + "</td>");

	var tcrwt = "<select name='tcrwt" + lstitem + "' id='tcrwt" + lstitem + "' class='lebartc ratatengah' onchange='getHaactual(" + lstitem + ")'></select>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + tcrwt + "</td>");

	var haactual = "<input type='text' name='haactual" + lstitem + "' id='haactual" + lstitem + "' style='background-color:#eee;' value='' readonly class='lebarhaactual ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + haactual + "</td>");

	var rate = "<input type='text' name='rate" + lstitem + "' id='rate" + lstitem + "' value='0' onkeyup='hanyaAngka(this)' style='background-color:#eee;' class='lebarhaactual ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + rate + "</td>");
	
	var mandayform = "<select name='mandayform" + lstitem + "' id='mandayform" + lstitem + "' class='lebarnormahk ratatengah'></select>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + mandayform + "</td>");
	
	var luasdikerjakan = "<input type='text' name='luasdikerjakan" + lstitem + "' id='luasdikerjakan" + lstitem + "' value='0' onkeyup='hanyaAngka(this);getTotalhk(" + lstitem + ");' class='lebarhaactual ratakanan amtluasdikerjakan'>";
	$('#trb' + lstitem).append("<td class='ratakanan'>" + luasdikerjakan + "</td>");

	var ttlhk = "<input type='text' name='ttlhk" + lstitem + "' id='ttlhk" + lstitem + "' value='0' onkeyup='hanyaAngka(this);' style='background-color:#eee;'  readonly class='lebarhaactual ratakanan amtttlhk'>";
	$('#trb' + lstitem).append("<td class='ratakanan'>" + ttlhk + "</td>");
	
	var rotasi = "<input type='text' name='rotasi" + lstitem + "' id='rotasi" + lstitem + "' value='0' onkeyup='hanyaAngka(this);getAmount(" + lstitem + ");' class='lebartc ratakanan amtrotasi'>";
	$('#trb' + lstitem).append("<td class='ratakanan'>" + rotasi + "</td>");

	var amount = "<input type='text' name='amount" + lstitem + "' id='amount" + lstitem + "' style='background-color:#eee;' value='0' onkeyup='hanyaAngka(this)' class='lebarlokasi ratakanan'><input type='hidden' class='amthk' name='amountbck" + lstitem + "' id='amountbck" + lstitem + "' value='0'>";
	$('#trb' + lstitem).append("<td class='ratakanan'>" + amount + "</td>");

	var rateborongan = "<input type='text' name='rateborongan" + lstitem + "' id='rateborongan" + lstitem + "' value='0' onkeyup='hanyaAngka(this);getAmountborongan(" + lstitem + ");' class='lebartc ratatengah'>";
	$('#trb' + lstitem).append("<td class='ratakanan'>" + rateborongan + "</td>");

	var amountborongan = "<input type='text' name='amountborongan" + lstitem + "' id='amountborongan" + lstitem + "' style='background-color:#eee;' value='0' onkeyup='hanyaAngka(this)' class='lebarlokasi ratakanan'><input class='amtborongan' type='hidden' name='amountboronganbck" + lstitem + "' id='amountboronganbck" + lstitem + "' value='0'>";
	$('#trb' + lstitem).append("<td class='ratakiri'>" + amountborongan + "</td>");

	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowPerawatan(" + lstitem + ");' />";
	$('#trb' + lstitem).append("<td>" + tambahbutton + "</td>");

	$('#trb' + lstitem).append("<td>" + kurangbutton + "</td>");

	$('#jmlrawat').val(lstitem);

	showActivityrwt("activityrwt" + lstitem);

	
}

//hapus row rkbperawatan
function deleterowBahan(i) {
	Spry.Widget.Utils.destroyWidgets('activityrwt' + i);
	Spry.Widget.Utils.destroyWidgets('tcrwt' + i);
	Spry.Widget.Utils.destroyWidgets('haactual' + i);
	Spry.Widget.Utils.destroyWidgets('rate' + i);
	Spry.Widget.Utils.destroyWidgets('mandayform' + i);
	Spry.Widget.Utils.destroyWidgets('luasdikerjakan' + i);
	Spry.Widget.Utils.destroyWidgets('rotasi' + i);
	Spry.Widget.Utils.destroyWidgets('rateborongan' + i);
	Spry.Widget.Utils.destroyWidgets('amount' + i);
	Spry.Widget.Utils.destroyWidgets('amountborongan' + i);
	 $('#trb' + i).remove();
	// getjmlRowbhn();
}

//get activity item
function showActivityrwt(selid) {
	
	$('#activityrwt').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp?panen=' + 5,
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
	var tahun = $('#tahun').val();
	var bulan = $('#bulan').val();
	var sectionname = $('#sectionname').val();
	var activityrwt = $('#activityrwt' + i).val();
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

function getHaactual(i)
{
	var tcrwt = $('#tcrwt' + i).val();
	if (tcrwt != "") {
		$.post("bkm_getactivity.asp?tcrwt=" + tcrwt + "&panen=" +  6, function(data) {
			$('#haactual' + i).empty();
			var json  = JSON.parse(data);
			var luas_gross = json.results[0].luas_gross
			$('#haactual' + i).val(luas_gross);
						
		});
		
	}
}

function getRate(i)
{
	var activityrwt = $('#activityrwt' + i).val();
	if (activityrwt != "") {
		$.post("bkm_getactivity.asp?panen=" +  7, function(data) {
			$('#haactual' + i).empty();
			var json  = JSON.parse(data);
			var rate_1 = json.results[0].rate_1
			$('#rate' + i).val(rate_1);
						
		});
		
	}
}

//tampilkan norma mandays
function getNormamandays(i)
{
	var activityrwt = $('#activityrwt' + i).val();
	$('#mandayform' + i).empty();
	$('#mandayform' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp?activity=' + activityrwt + "&panen=" + 8,
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

function getTotalhk(i)
{
	var mandayform 		= parseFloat($("#mandayform" + i).val());
	var luasdikerjakan 	= parseFloat($("#luasdikerjakan" + i).val());
	var ttlhkrp			= mandayform*luasdikerjakan;
	$("#ttlhk" + i).val(ttlhkrp);
	
}

function getAmount(i)
{ 
	var rate 			= parseFloat($("#rate" + i).val());
	var ttlhk 			= parseFloat($("#ttlhk" + i).val());
	var rotasi 			= parseFloat($("#rotasi" + i).val());
	var amountttl 		= (rate*ttlhk)*rotasi;
	amountttl			= Math.round(amountttl);
	$("#amount" + i).val(numformat(amountttl));
	$("#amountbck" + i).val((amountttl));
}

function getAmountborongan(i)
{ 
	var rateborongan 			= parseFloat($("#rateborongan" + i).val());
	var luasdikerjakan 			= parseFloat($("#luasdikerjakan" + i).val());
	var amountttl 		= (rateborongan*luasdikerjakan);
	amountttl			= Math.round(amountttl);
	$("#amountborongan" + i).val(numformat(amountttl));
	$("#amountboronganbck" + i).val((amountttl));
}

function getPuttotal(i) {
		var jmlluasdikerjakan	= 0;
		var jmlttlhk			= 0;
		var jmlttlrotasi		= 0;
		var jmlttlamount		= 0;
		var jmlrawat			= parseInt($('#jmlrawat').val());
		for(var i=0; i<jmlrawat; i++) {
			 jmlluasdikerjakan	    += ($('#luasdikerjakan' + i).val());
			// jmlttlhk	    		+= parseInt($('#ttlhk'+i).val());
			// jmlttlrotasi	    	+= parseInt($('#rotasi'+i).val());
			jmlttlamount	    	+= parseInt($('#amountbck'+ i).val());
		  	alert(jmlluasdikerjakan);
			
		}	
		
		// $('#amountluasdikerjakanfrm').val((jmlluasdikerjakan));
		// $('#amountttlhkfrm').val((jmlttlhk));
		// $('#amountrotasifrm').val((jmlttlrotasi));
		$('#amountrkbrpfrm').val((jmlttlamount));
		
}

var StartIterationitem=1;
function addrowBahan(trid)
{

	var lstitem = this.StartIterationitem++;
	
	if (lstitem < 2) {
		var row = $("#trheader2").after("<tr id='trb1" + lstitem + "'></tr>");
		var kurangbutton = "&nbsp;";
	} else {
		var row = $('#trb1' + trid).after("<tr id='trb1" + lstitem + "'></tr>");
		var kurangbutton = "<img src='css/login/images/del.png' onclick='deleterowBahan(" + lstitem + ");' />";
	}

	var activityitem = "<select name='activityitem" + lstitem + "' id='activityitem" + lstitem + "' class='lebaractbarang1' onchange='getTCitem(" + lstitem + ");showItem(" + lstitem + ")'></select>";
	$('#trb1' + lstitem).append("<td class='ratakiri'>" + activityitem + "</td>");

	var tcitem = "<select name='tcitem" + lstitem + "' id='tcitem" + lstitem + "' class='lebartc ratatengah'></select>";
	$('#trb1' + lstitem).append("<td class='ratakiri'>" + tcitem + "</td>");

	var itemcode = "<select name='itemcode" + lstitem + "' id='itemcode" + lstitem + "' onchange='getSatuan(" + lstitem + ")' class='lebaritembarang'></select>";
	$('#trb1' + lstitem).append("<td class='ratakiri'>" + itemcode + "</td>");

	var uom = "<input type='text' name='uom" + lstitem + "' id='uom" + lstitem + "' value='' readonly class='lebaruom ratatengah'>";
	$('#trb1' + lstitem).append("<td class='ratakiri'>" + uom + "</td>");

	var normamm = "<input type='text' name='normamm" + lstitem + "' id='normamm" + lstitem + "' value='0' class='lebaruom ratatengah'>";
	$('#trb1' + lstitem).append("<td class='ratakiri'>" + normamm + "</td>");

	var qtyitem = "<input type='text' name='qtyitem" + lstitem + "' id='qtyitem" + lstitem + "' value='0' onkeyup='hanyaAngka(this);getAmountitem(" + lstitem + ")' class='lebaruom ratatengah amtqty'>";
	$('#trb1' + lstitem).append("<td class='ratakanan'>" + qtyitem + "</td>");

	var rateitem = "<input type='text' name='rateitem" + lstitem + "' id='rateitem" + lstitem + "' value='0' onkeyup='hanyaAngka(this);getAmountitem(" + lstitem + ")' class='lebarnormahk ratatengah'>";
	$('#trb1' + lstitem).append("<td class='ratakanan'>" + rateitem + "</td>");

	var amountitem = "<input type='text' name='amountitemfrm" + lstitem + "' id='amountitemfrm" + lstitem + "' readonly value='0' onkeyup='hanyaAngka(this)' class='lebarnormahk ratatengah'><input type='hidden' name='amountitemx" + lstitem + "' id='amountitemx" + lstitem + "' readonly value='0' onkeyup='hanyaAngka(this)' class='lebarnormahk ratatengah amtbahan'>";
	$('#trb1' + lstitem).append("<td class='ratakanan'>" + amountitem + "</td>");

	var tambahbutton = "<img src='css/login/images/add.jpg' onclick='addrowBahan(" + lstitem + ");' />";
	$('#trb1' + lstitem).append("<td>" + tambahbutton + "</td>");

	$('#trb1' + lstitem).append("<td>" + kurangbutton + "</td>");

	$('#jumldataitem').val(lstitem);


	showActivityitem("activityitem" + lstitem);
	new Spry.Widget.ValidationTextField("normamm" + lstitem);
	new Spry.Widget.ValidationTextField("qtyitem" + lstitem);
	new Spry.Widget.ValidationTextField("rateitem" + lstitem);
	
}

//hapus row bahan
function deleterowBahan(i) {
	Spry.Widget.Utils.destroyWidgets('itemcode' + i);
	Spry.Widget.Utils.destroyWidgets('uom' + i);
	Spry.Widget.Utils.destroyWidgets('activityitem' + i);
	Spry.Widget.Utils.destroyWidgets('tcitem' + i);
	Spry.Widget.Utils.destroyWidgets('normamm' + i);
	Spry.Widget.Utils.destroyWidgets('qtyitem' + i);
	Spry.Widget.Utils.destroyWidgets('rateitem' + i);
	Spry.Widget.Utils.destroyWidgets('amountitem' + i);
	Spry.Widget.Utils.destroyWidgets('amountitemfrm' + i);
	 $('#trb1' + i).remove();
	// getjmlRowbhn();
}

//get activity item
function showActivityitem(selid) {
	$('#activityitem').empty();
	$('#' + selid).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getactivity.asp?panen='+14,
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
function getTCitem(i)
{
	var tahun = $('#tahun').val();
	var bulan = $('#bulan').val();
	var sectionname = $('#sectionname').val();
	var activityitem = $('#activityitem' + i).val();
	$('#tcitem' + i).empty();
	$('#tcitem' + i).select2({
		ajax: {
			delay: 1000,
			url : 'bkm_getblok_rkb.asp?activity=' + activityitem + "&tcby=" + 1 + "&tahun=" + tahun + "&bulan=" + bulan + "&section=" + sectionname,
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

//tampilkan item
function showItem(i) {
	var activityitem = $('#activityitem' + i).val();
	var fcbatxt 	= $("#fcbatxt").val();
	$('#itemcode' + i).empty();
	$('#itemcode' + i).select2({
		ajax: {
			delay: 1000,
			url : "bkm_getactivity.asp?panen=" + 15 + "&activity=" + activityitem + "&fcba=" + fcbatxt,
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
	var itemcode = $('#itemcode' + i).val();
	var jobcode = $('#activityitem' + i).val();
	if (itemcode != "") {
		$.post("bkm_employee.asp?itemcode=" + itemcode +"&satuan=" + 8 + "&jobcode=" + jobcode, function(data) {
			$('#uom' + i).empty();
			var json  		= JSON.parse(data);
			var uom 		= json.results[0].uomactivematerial
			var actmatnorm 	= json.results[0].actmatnorm
			$('#uom' + i).val(uom);
			$('#normamm' + i).val(actmatnorm);
						
		});
		
	}
}

function getAmountitem(i)
{
	var qtyitem 		= parseFloat($('#qtyitem' + i).val());
	var rateitem 		= parseFloat($('#rateitem' + i).val());
	if(qtyitem==""){
		qtyitem=0;
	}
	if(rateitem==""){
		rateitem=0;
	}
	var amountitem  	= qtyitem*rateitem;
	var amountitemfrm	= numformat(amountitem);
	$('#amountitemfrm' + i).val(amountitemfrm);
	$('#amountitemx' + i).val(amountitem);
}

var _hitungLuasdikerjakan = () => {
    let tgrandtotal = 0;
    $('.amtluasdikerjakan').each(function(index,element) {
      tgrandtotal += Number($("input[name^='luasdikerjakan']").eq(index).val());     
    });  
  
    tgrandtotal = tgrandtotal.toString().replace('.',',');            
  
    if(tgrandtotal==0) tgrandtotal='0,00';

	let grandtotal = numformat(tgrandtotal);
    $('#amountluasdikerjakan').val(tgrandtotal).attr('placeholder',tgrandtotal);      
    $('#amountluasdikerjakanfrm').val(grandtotal).attr('placeholder',grandtotal);      
    return;
}

var _hitungTtlhk = () => {
    let tgrandtotal = 0;
    $('.amtttlhk').each(function(index,element) {
      tgrandtotal += Number($("input[name^='ttlhk']").eq(index).val());     
    });  
  
    tgrandtotal = tgrandtotal.toString().replace('.',',');            
  
    if(tgrandtotal==0) tgrandtotal='0,00';

	let grandtotal = numformat(tgrandtotal);
    $('#amountttlhk').val(tgrandtotal).attr('placeholder',tgrandtotal);      
    $('#amountttlhkfrm').val(grandtotal).attr('placeholder',grandtotal);      
    return;
}

var _hitungTtlrotasi = () => {
    let tgrandtotal = 0;
    $('.amtrotasi').each(function(index,element) {
      tgrandtotal += Number($("input[name^='rotasi']").eq(index).val());     
    });  
  
    tgrandtotal = tgrandtotal.toString().replace('.',',');            
  
    if(tgrandtotal==0) tgrandtotal='0,00';

	let grandtotal = numformat(tgrandtotal);
    $('#amountrotasi').val(tgrandtotal).attr('placeholder',tgrandtotal);      
    $('#amountrotasifrm').val(grandtotal).attr('placeholder',grandtotal);      
    return;
}

var _hitungTtlAmount = () => {
    let tgrandtotal = 0;
    $('.amthk').each(function(index,element) {
      tgrandtotal += Number($("input[name^='amountbck']").eq(index).val());     
    });  
  
    tgrandtotal = tgrandtotal.toString().replace('.',',');            
  
    if(tgrandtotal==0) tgrandtotal='0,00';

	let grandtotal = numformat(tgrandtotal);
    $('#amountrkbrp').val(tgrandtotal).attr('placeholder',tgrandtotal);      
    $('#amountrkbrpfrm').val(grandtotal).attr('placeholder',grandtotal);      
    return;
}

var _hitungAmountBorongan = () => {
    let tgrandtotal = 0;
    $('.amtborongan').each(function(index,element) {
      tgrandtotal += Number($("input[name^='amountboronganbck']").eq(index).val());     
    });  
  
    tgrandtotal = tgrandtotal.toString().replace('.',',');            
  
    if(tgrandtotal==0) tgrandtotal='0,00';

	let grandtotal = numformat(tgrandtotal);
    $('#amountborrkbrp').val(tgrandtotal).attr('placeholder',tgrandtotal);      
    $('#amountborrkbrpfrm').val(grandtotal).attr('placeholder',grandtotal);      
    return;
}

//hitung total qty bahan
var _hitungQtyBahan = () => {
    let tgrandtotal = 0;
    $('.amtqty').each(function(index,element) {
      tgrandtotal += Number($("input[name^='qtyitem']").eq(index).val());     
    });  
  
    tgrandtotal = tgrandtotal.toString().replace('.',',');            
  
    if(tgrandtotal==0) tgrandtotal='0,00';

	let grandtotal = numformat(tgrandtotal);
    $('#ttlqty').val(tgrandtotal).attr('placeholder',tgrandtotal);      
    $('#ttlqtyfrm').val(grandtotal).attr('placeholder',grandtotal);      
    return;
}

var _hitungAmountBahan = () => {
    let tgrandtotal = 0;
    $('.amtbahan').each(function(index,element) {
      tgrandtotal += Number($("input[name^='amountitemx']").eq(index).val());     
	  //alert(tgrandtotal);
    });  
  
    tgrandtotal = tgrandtotal.toString().replace('.',',');            
  
    if(tgrandtotal==0) tgrandtotal='0,00';

	let grandtotal = numformat(tgrandtotal);
    $('#amountbahan').val(tgrandtotal).attr('placeholder',tgrandtotal);      
    $('#amountbahanfrm').val(grandtotal).attr('placeholder',grandtotal);      
    return;
}

 </script>