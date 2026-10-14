<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include file="mm.inc"-->
<!--#include file="DailyValidation.inc"-->
<!--#include file="qtyOnHandFunction.inc"-->
<%
dim businessunit,fcbaname,CompanyName,q,sqlhdr,sqldtl,simpan
dim returnissueno,tglreturn,store,noref,remarkhdr
dim nQtdFields,nCount
dim itemcode,itemname,uom,qtyissue,qtyreturn,remarkdtl,linenokey
dim sqlnomor,rsnomor,tglprev,currentnomor,currentnomorfix
dim bulansplit,tglsplit,thnsplit
dim qTotalRecordOnHand,rsTotalRecordOnHand,tempTotalRecordOnHand,Quantity
dim qryUpdatestockOnhand
dim ScriptName,modulfile,executeQry,rsStockLedger

businessunit= trim(session("fcba"))
simpan		= trim(request("simpan"))
store		= trim(request("store"))
tglreturn	= trim(request("tglreturn"))
store		= trim(request("store"))
noref		= trim(request("noref"))
remarkhdr	= trim(request("remarkhdr"))

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
 
end if

if simpan = "SIMPAN" THEN 
	bulansplit  = left(tglreturn,2)
	tglsplit    = mid(tglreturn,4,2)
	thnsplit    = right(tglreturn,4)
	' buat nomor transaksi'
	sqlnomor = "SELECT SUBSTR(returnno,-5,5) AS LastNo " & _
	" FROM RETURNISSUEHDR " & _
	" WHERE SUBSTR(returnno,-10,4)='" & thnsplit & "' " & _
	" AND FCBA='" & businessunit & "' ORDER BY SUBSTR(returnno,-5,5) DESC FETCH FIRST 1 ROWS ONLY"
	Set rsnomor = Server.CreateObject("ADODB.Recordset")
	rsnomor.CursorType = 3
	rsnomor.CursorLocation = 3
	rsnomor.Open sqlnomor, connstr
	if rsnomor.eof then
			tglprev = "0000"
			currentnomor = tglprev + 1
			currentnomor = "0000" & currentnomor
			currentnomorfix = right(currentnomor,5)
			returnissueno=businessunit&"/"&"RT"&"/"&bulansplit&thnsplit&"/"&currentnomorfix
			'Response.Write "ok"
	else 
			tglprev = rsnomor("LastNo")
			currentnomor = tglprev + 1
			currentnomor = "0000" & currentnomor
			currentnomorfix = right(currentnomor,5)
			returnissueno	= businessunit&"/"&"RT"&"/"&bulansplit&thnsplit&"/"&currentnomorfix
	end if
	rsnomor.close 
	set rsnomor = nothing
'	returnissueno = 1
	'simpan ke tabel hdr return
	sqlhdr = "insert into RETURNISSUEHDR(RETURNNO, RETURNDATE, STORE, NOREF, REMARK, " & _
	" FCENTRY,FCEDIT,FCIP,FCBA,LASTUPDATE,LASTTIME,POSTINGFLG) " & _
	" values('" & returnissueno & "',TO_DATE('" & tglreturn & "', 'mm/dd/yyyy'),'" & store & "','" & noref & "','" & remarkhdr & "', " & _ 
	" '" & Session("CodeName") & "', " & _ 
	" '" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & businessunit & "',SYSDATE,SYSDATE,'0')"
	ExecuteAndLog "StockReturn.asp", sqlhdr

	'connstr.BeginTrans
	'connStr.execute(sqlhdr)

	'looping di tabel detail
	nQtdFields 	= request.form("jumldata")
	nCount 		= 1

	do while cint(nCount) <= cint(nQtdFields) 
		itemcode 		= request.form("itemcode" & nCount)
	'	fcname 			= request.form("fcname" & nCount)
		uom 			= request.form("uom" & nCount)
		qtyissue 		= request.form("qtyissue" & nCount)
		qtyreturn 		= request.form("qtyreturn" & nCount)
		remarkdtl 			= request.form("remark" & nCount)
		linenokey 			= request.form("linenokey" & nCount)
		if qtyreturn > 0 then

			sqldtl = "insert into RETURNISSUEDTL(RETURNNO, ITEMCODE, UOM, QTYISSUE, QTYRETURN,REMARKDTL,LINENOKEY,FCBA) " & _
			" values('" & returnissueno & "','" & itemcode & "','" & uom & "','" & qtyissue & "', " & _ 
			" '" & qtyreturn & "','" & remarkdtl & "','" & linenokey & "','" & businessunit & "')"
			ExecuteAndLog "StockReturn.asp", sqldtl
			'connstr.BeginTrans
			'connStr.execute(sqldtl)

			' modify jun (19 juni 2024)
			qTotalRecordOnHand  =   "select count(*) totalRecordOnHand from stockonhand " & _
									"where " & _
									"storecode = '" & store & "' and " & _												 
									"stockcode = '" & ITEMCODE & "' and " & _
									"fcba = '" &  businessunit & "'"
			'Response.Write qTotalRecordOnHand
			set rsTotalRecordOnHand = Server.CreateObject("ADODB.Recordset")
			rsTotalRecordOnHand.Open qTotalRecordOnHand, connstr
			if not rsTotalRecordOnHand.eof then
					tempTotalRecordOnHand = cdbl(rsTotalRecordOnHand("totalRecordOnHand"))
			else
					tempTotalRecordOnHand = 0
			end if 
				
			If tempTotalRecordOnHand = 0 then
				
					' Tambahkan item stock dan qty di tabel stock onhand
					qryinsertstockOnhand 	= cmdInsertStockOnHand(store,ITEMCODE,qtyreturn,"RT")
					
			Else	 
					' Update qty di tabel stock onhand
					qryUpdatestockOnhand 	= cmdUpdateStockOnHand(store,ITEMCODE,qtyreturn,"RT","Insert")
			End If	

			'CEK ITEM DAHULU DI TABEL STOCKLEDGER
			'JIKA TIDAK ADA INSERT BARU, SELAIN ITU UPDATE
			dim qryStockLedger, temp,totalRecordStockLedger,qryinsertstockledger,qryupdateStockLedger
			qryStockLedger = "SELECT COUNT(*) totalRecordStockLedger FROM STOCKLEDGER" & _ 
			" WHERE " & _ 
			" STORECODE = '" & store & "' and " & _
			" ITEMCODE  = '" & ITEMCODE & "' and " & _ 
			" year  	= to_char(to_date('" & tglreturn & "','mm/dd/yyyy'),'yyyy') and " & _ 
			" month 	= to_char(to_date('" & tglreturn & "','mm/dd/yyyy'),'mm') and " & _ 
			" FCBA 		= '" & businessunit & "' "

			'Response.Write qryStockLedger
			'Response.end
			
			set rsStockLedger = Server.CreateObject("ADODB.Recordset")
			rsStockLedger.Open qryStockLedger, connstr
			if not rsStockLedger.eof then
					totalRecordStockLedger = cdbl(rsStockLedger("totalRecordStockLedger"))
			else
					totalRecordStockLedger = 0
			end if 
		

			'function ini digunakan untuk insert/update data di tabel stockledger
			' modify jun (19 juni 2024)
			If totalRecordStockLedger = 0 then
			
				'insert stockledger
				qryinsertstockledger 	= cmdInsertStockLedger(store,ITEMCODE,tglreturn,qtyreturn, "RT")
			ELSE
			
				'update stockledger
				qryupdateStockLedger 	= cmdUpdateStockLedger(store,ITEMCODE,tglreturn,qtyreturn, "RT", "Insert")
			END IF
		'	Response.Write qryupdateStockLedger
		'	Response.end

		end if

	 nCount 	= nCount + 1
	loop
	
END IF
%>

<html>
	<head>
		<title>FBPlants - Stock Return</title>
		<script type="text/javascript" language="JavaScript1.2" src="stlib.js" ></script>
		<link rel="StyleSheet" href="dtree.css" type="text/css" />
		<link rel="StyleSheet" href="style-content.css" type="text/css" />
		<link rel="stylesheet" href="css/jquery-ui.css" />
		<!-- HADIR/ font ICON -->
		<link rel="stylesheet" href="css/font/font-awesome.css" />
		<!-- END:HADIR -->
		<script type="text/javascript" src="dtree.js"></script>
		<script src="/js/jquery-1.9.1.js"></script>
		<script src="/js/jquery-ui.js"></script>

		<link href="/js/tabbedpanels/SpryTabbedPanels.css" rel="stylesheet" type="text/css"/>
		<script src="/js/tabbedpanels/SpryTabbedPanels.js" type="text/javascript"></script>
		
		
		<link href="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.css" rel="stylesheet" type="text/css"/>
		<script src="/js/spry/widgets/textfieldvalidation/SpryValidationTextField.js" type="text/javascript"></script>
		
		<link href="/js/spry/widgets/selectvalidation/SpryValidationSelect.css" rel="stylesheet" type="text/css"/>
		<script src="/js/spry/widgets/selectvalidation/SpryValidationSelect.js" type="text/javascript"></script>

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
	.posisikiri{
		position: relative;
		left: -15px;
	}
	.lebarsection{
		width: 220px;
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

<body<% if Request("oldui") <> "1" then %> class="dash page-show page-legacy page-doc"<% end if %>>
<!--#include file="top.jsp"-->
<div style=" width:21.7%; float:left;margin-left:0.3%; margin-top:0.3%; margin-bottom:0.3%;">
 <!--#include file="menu.asp"-->
</div>
<div class="judul_head" valign="top" bgcolor="#fff" style=" width:77.4%; margin-left:0.3%; margin-right:0.3%; margin-top:0.3%; margin-bottom:0.3%; float:left; border:none;"> 
 	<div class="judul">
	   Module : <%=request("txtModule")%> / <%=request("txtTable")%>
	</div>
<div class="box_both" style="overflow:scroll;">  

<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-undo" aria-hidden="true"></i>Header Stock Return</div><% end if %>
<form name="frmEntry" method="post" action="stockreturn.asp">
	<input type="hidden" name="jumldata" id="jumldata" value="0" />
	<input type="hidden" name="proses" id="proses" value="1" />
	<table class="data" width="100%">	
		<tr>
		  <td align="left">Order No</td>
		  <td align="left"><input type="text" maxlength="20" name="OrderNo" size="30" value="(Auto)" style="background-color:#eee;"  readonly autocomplete = "Off">
		  </td>
		</tr>
		<tr>
		  <td align="left">Date.Return</td>
		  <td align="left">
		  <input type="text" class="datepicker" maxlength="15" name="tglreturn" readonly id="tglreturn" size="15" value="<%=textDate(date)%>" title="Tanggal Return">
		  <font color="#FF0000">&nbsp;*</font></td>
		</tr>	

		<tr>
			<td align="left" >No.Issue</td>
			<td align="left"><input type="text" readonly name="noref" id="noref" onkeyup="getDetailreturn();" >
			<font color="#FF0000">&nbsp;*</font> 
			&nbsp;&nbsp;<input type="button" onclick="Popupissue()" value="Cari&nbsp;&nbsp;..."></input>
			</td>
		</tr>
		<tr>
			<td align="left" >Store</td>
			<td align="left">
				<select name="store" id="store" class="lebargangcode" style=height:24px>
					<option value="ST1">ST1 - GUDANG KEBUN</option> 
				</select>
			</td>
		</tr>	
		<tr>
			<td align="left" >Remark</td>
			<td align="left">
				<textarea name="remarkhdr" id="remarkhdr" rows="6" cols="40">
				</textarea>
			</td>			
		</tr>

	</table>
	<% if Request("oldui") <> "1" then %><div class="fbp-doc-title"><i class="fa fa-list" aria-hidden="true"></i>Item yang Dikembalikan<small>pilih No. Issue untuk memuat item</small></div><% end if %>
	<table class="data" cellpadding="0" cellspacing="0" width="100%" id='tableall'>
		<tr>
			<td >
			<table class="data" id='table2'> 
						<tr id="trheader2">
							<th>Item Code</th>
							<th>Item Name</th>
							<th>UOM</th>
							<th>Qty Issue</th>
							<th>Qty Return</th>
							<th colspan="1">Remark</th>
						</tr>
						<tr id='tru'>
							<tbody id="bodytable">
							</tbody>
						</tr>

						<tr>
							<td colspan='6' align=center><input Type="Submit"  name="simpan" value="SIMPAN">
							

							</td>
						
						</tr>
						
					</table>
			</td>
		</tr>
		
	</table>

</form>
</div></div>
  <!--#include file="bottom.jsp"-->
</body>
</html>

<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
<script src="/js/select2/select2.min.js" type="text/javascript"></script>

<script type="text/javascript">

	function hanyaAngka(evt) {
		evt.value = evt.value.replace(/[^0-9,.-]/g,"");
	}

$(function() {	
	$("#tglreturn").datepicker();


});

function Popupissue() {
	var newwin   = window.open("stockissuepopup.Asp?Mode=Print","newwindow","width = 700, height = 505, screenX = 0, screenY = 0, directories = 0, fullscreen = 0, location = 0, menubar = 0, scrollbars = 1, status = 0, toolbar = 1");
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

//getEmployee
function getDetailreturn() {
	var noref = $("#noref").val();
	if (noref != "") {
		$.post("bkm_viewdata.asp?noref=" + noref + "&idview=" + 28, function(data) {
			$('#bodytable').empty();
			var json  = JSON.parse(data);
			var jumlview  = json.results.length;
			var no = 1;
			for (var i = 0; i < jumlview; i++) {
				var sivcode 	= json.results[i].sivcode;
				var stockcode 	= json.results[i].stockcode;
				var fcname 		= json.results[i].fcname;
				var unitofmeasurement 		= json.results[i].unitofmeasurement;
				var quantity 	= json.results[i].quantity;
				var remarks 	= json.results[i].remarks;
				var linenokey 	= json.results[i].linenokey;
				
				var docketlist = '';
					docketlist += '<tr id="tru'+no+'">';
					docketlist += '<td align="center">' + stockcode + '<input type="hidden" class="ratatengah lebarhk" name="itemcode' + no + '" id="itemcode' + no + '" value="' + stockcode + '"/><input type="hidden" class="ratatengah lebarhk" name="linenokey' + no + '" id="linenokey' + no + '" value="' + linenokey + '"/></td>';
					docketlist += '<td align="center">' + fcname + '<input type="hidden" class="ratatengah lebarhk" name="fcname' + no + '" id="fcname' + no + '" value="' + fcname + '"/></td>';	
					docketlist += '<td align="center">' + unitofmeasurement + '<input class="ratatengah lebarhk" type="hidden" name="uom' + no + '" id="uom' + no + '" value="' + unitofmeasurement + '"/></td>';
					docketlist += '<td align="center">' + quantity + '<input class="ratatengah lebarabsen" type="hidden" name="qtyissue' + no + '" id="qtyissue' + no + '" value="' + quantity + '"/></td>';
					docketlist += '<td align="center"><input class="ratatengah lebarabsen" type="text" onkeyup="hanyaAngka(this)" name="qtyreturn' + no + '" id="qtyreturn' + no + '" value="0"/></td>';
					docketlist += '<td align="center"><textarea rows="3" cols="20" name="remark' + no + '" id="remark' + no + '">' + remarks + '</textarea></td>';		
					docketlist += '</tr>';
					$('#bodytable').append(docketlist);
				new Spry.Widget.ValidationTextField("qtyreturn" + no);
				no++;
				
			}
			$('#jumldata').val(jumlview);
			
		});
		
	} else {
		$('#bodytable').empty();
	}
	
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


</script>