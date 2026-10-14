<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim businessunit,fcbaname,CompanyName,q,sqltotdata,bkmno 
dim rsCompany,company
bkmno= trim(request("rkbno"))
businessunit= trim(session("fcba"))

q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
	' response.write q
set rsCat = connStr.Execute(q)
if not rsCat.eof then
   fcbaname = rsCat("fcName")
   CompanyName=rsCat("FCCOMPANYNAME")
 
end if


' get data header
dim sqlheader,rsHeader,bkmnohdr,DATECREATED,SECTIONNAME,GANGCODE,ROWSTATEID
dim period_month,period_year 

sqlheader = "select rkbno,period_month,period_year,sectioncode,document_classification,rowstate,totupah,totbahan,totvra,totwip,classification_api.get_fcname(document_classification) docname " & _
            " from rkbhdr where rkbno='" & bkmno & "' AND fcba='" & businessunit & "'"
set rsHeader = connStr.Execute(sqlheader)

if not rsHeader.eof then 
	bkmnohdr = rsHeader("rkbno")
	period_month = rsHeader("period_month")
	period_year = rsHeader("period_year")
	SECTIONNAME = rsHeader("sectioncode")
	Clasification = rsHeader("document_classification")
	Clasificationname = rsHeader("docname")
	
	ROWSTATE = rsHeader("rowstate")
end if
DIM simpan,tglinput,sqlhdr,tahun,bulan,contractorcode,nokontrak,amountrkbrp,jmlrkb,nQtdFields,nCount
DIM actcode,blok,qtyrkb,hargarkb,linenokey,sqldtl,fcbakontrak
dim bulansplit,tglsplit,thnsplit,sqlnomor,rsnomor,tglprev,currentnomor,currentnomorfix,rkbno,chargeunit
simpan		= trim(request("simpan"))
chargeunit	= trim(request("chargeunit"))
tglinput	= trim(request("tglinput"))
tahun		= trim(request("tahun"))
bulan		= trim(request("bulan"))
contractorcode		= trim(request("contractorcode"))
nokontrak		= trim(request("nokontrak"))
amountrkbrp		= trim(request("amountrkbrp"))
fcbakontrak		= trim(request("businessunitar"))
if simpan = "SAVE" THEN 
	bulansplit  = left(tglinput,2)
	tglsplit    = mid(tglinput,4,2)
	thnsplit    = right(tglinput,4)
	' buat nomor transaksi'
	sqlnomor = "SELECT SUBSTR(RKBNO,-5,5) AS LastNo " & _
	" FROM RKBHDR " & _
	" WHERE SUBSTR(RKBNO,-11,4)='" & thnsplit & "' " & _
	" AND FCBA='" & chargeunit & "' AND DOCUMENT_CLASSIFICATION='504' ORDER BY SUBSTR(RKBNO,-5,5) DESC FETCH FIRST 1 ROWS ONLY"
	Set rsnomor = Server.CreateObject("ADODB.Recordset")
	rsnomor.CursorType = 3
	rsnomor.CursorLocation = 3
	rsnomor.Open sqlnomor, connstr
	if rsnomor.eof then
			tglprev = "0000"
			currentnomor = tglprev + 1
			currentnomor = "0000" & currentnomor
			currentnomorfix = right(currentnomor,5)
			rkbno=businessunit&"/"&"RKB"&"/"&bulansplit&thnsplit&"/"&currentnomorfix
			'Response.Write "ok"
	else 
			tglprev = rsnomor("LastNo")
			currentnomor = tglprev + 1
			currentnomor = "0000" & currentnomor
			currentnomorfix = right(currentnomor,5)
			rkbno	= businessunit&"/"&"RKB"&"/"&bulansplit&thnsplit&"/"&currentnomorfix
	end if
	rsnomor.close 
	set rsnomor = nothing

	sqlhdr = "insert into RKBHDR(RKBNO, PERIOD_YEAR, PERIOD_MONTH, SECTIONCODE, CREATEDATE, " & _
	" LASTUPDATE,FCENTRY,FCEDIT,FCIP,FCBA,ROWSTATE,DOCUMENT_CLASSIFICATION,TOTWIP,NOKONTRAK,CHARGEUNIT) " & _
	" values('" & rkbno & "','" & tahun & "','" & bulan & "','" & contractorcode & "',TO_DATE('" & tglinput & "', 'dd/mm/yyyy'),SYSDATE," & _ 
	" '" & Session("CodeName") & "', " & _ 
	" '" & Session("CodeName") & "','" & Request.ServerVariables("REMOTE_ADDR") & "','" & fcbakontrak & "','Approved','504','" & amountrkbrp & "','" & nokontrak & "','" & chargeunit & "')"
	'connstr.BeginTrans
	'Response.Write sqlhdr
	connStr.execute(sqlhdr)

	'looping di tabel detail
	nQtdFields 	= request.form("jmlrkb")
	nCount 		= 1

	do while cint(nCount) <= cint(nQtdFields) 
		actcode 		= request.form("actcode" & nCount)
		blok 			= request.form("blok" & nCount)
		qtyrkb 			= request.form("rkbqty" & nCount)
		hargarkb 		= request.form("rate" & nCount)
		linenokey 			= request.form("lineno" & nCount)
		if qtyrkb > 0 then

			sqldtl = "insert into RKBWIPDTL(RKBNO, ACTCODE, LOCATIONCODE, QTY, HARGA,LINENOKEY,FCBA) " & _
			" values('" & rkbno & "','" & actcode & "','" & blok & "','" & qtyrkb & "', " & _ 
			" '" & hargarkb & "','" & linenokey & "','" & chargeunit & "')"
			'connstr.BeginTrans
			connStr.execute(sqldtl)

		end if

	 nCount 	= nCount + 1
	loop

END IF
%>

<html>
	<head>
		<title>FBPlants - RKB WIP</title>
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
		width: 48px;
	}
	.lebarlokasi{
		width: 60px;
	}
	.lebartc{
		width: 70px;
	}
	.lebarrate{
		width: 60px;
	}
	.lebaractivity{
		/*width: 135px;*/
	}
	.lebarallow{
		width: 130px;
	}
	.lebarupah{
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
	.lebaritembarang{
		width: 400px;
	}
	.lebaractbarang{
		width: 200px;
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



<body>
<!--#include file="top.jsp"-->

<br>
<br>
<br>
<br>
<br>


<div id="head_judul" style="float:left;font-size:10px; text-align:left;">
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : RKB WIP</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry" method="post" action="rkb_wip_input.asp">
	
	<table class="data posisikiri" border="1" width="80%" align="center">			
		<tr>
			<td  align="left">
				<font face="Verdana" size="1" align="LEFT">Company</font>
			</td>

			<td  align="left">

				<select name="companyar" id="companyar" onchange="getFCBAar()">
							<option value="0">--Pilih Company--</option> 
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
										<%if cstr(""&company)=cstr(""&rsCompany("companycode")) then %> selected 
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
			</td align="left">
				<td align="left" >FCBA Kontrak</td>
				<td align="left" colspan="3">
					<select name="businessunitar" id="businessunitar" class="lebaractbarang" onchange="getKontrak()">
						<option value="">--Pilih FCBA Kontrak--</option>
					</select>
			</td>

			
		</tr>
		
		<tr>
			<td align="left" >No Kontrak</td>
			<td align="left">
				<select name="nokontrak" id="nokontrak" class="lebaractbarang" onchange="viewData()">
					<option value="">--Pilih Kontrak--</option>
				</select>
			</td>
			<td align="left" >Contractor Name</td>
			<td align="left" colspan="3">
				<input type="text" name="contractorcode" readonly id="contractorcode" size="10" value="" style="background-color:#eee;">
				<input type="text" name="contractorname" readonly id="contractorname" size="26" value="" style="background-color:#eee;">
				
			</td>
					
		</tr>
		<tr>
			<td align="left" >Charge Unit</td>
			<td align="left">
				<select name="chargeunit" id="chargeunit" class="lebaractbarang">
						<option value="">--Pilih Charge Unit--</option>
					</select>
			</td>
			
			<td align="left" >Period</td>
			<td align="left">
			    <select name="tahun" id="tahun" style="width:100px">
				<%
					dim sqltahun7,rstahun7
					sqltahun7 = "select distinct to_number(accountingyear) tahun from accountingyear order by tahun desc"  
					Set rstahun7 = Server.CreateObject("ADODB.Recordset")
					rstahun7.CursorType = 3
					rstahun7.CursorLocation = 3
					rstahun7.Open sqltahun7, connstr
					While Not rstahun7.EOF 
						if cint(period_year)=cint(rstahun7("tahun")) then %>
						<option value="<%=rstahun7("tahun")%>" Selected ><%=rstahun7("tahun")%></option>
						<%else%>
						<option value="<%=rstahun7("tahun")%>"><%=rstahun7("tahun")%></option>
						<%end if%>					
					<%rstahun7.movenext
					wend 
					rstahun7.close 
					

				%>	
				</select>
				<select name="bulan" id="bulan" style="width:120px" onchange="getKontrak()">
				<%
					dim sqlbulan7,rsbulan7
					sqlbulan7 = "select distinct to_number(accountingperiod) bulan,description from accountingperiod order by bulan " 
					Set rsbulan7 = Server.CreateObject("ADODB.Recordset")
					rsbulan7.CursorType = 3
					rsbulan7.CursorLocation = 3
					rsbulan7.Open sqlbulan7, connstr
					While Not rsbulan7.EOF 
						if cint(period_month)=cint(rsbulan7("bulan")) then %>
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
			<td align="left" >Tanggal Input</td>
			<td align="left">
				<input type="text" name="tglinput" readonly id="tglinput" size="30" value="<%=date()%>" title="BKM No">
			</td>
		</tr>

	</table>
	
	<br>
		<input type="hidden" name="jmlrkb" id="jmlrkb" value="0" />
	<table align=center border="0" class='data posisikiri' width='80%' id='tableall'>
		<tr>
			<td>
			   		<div class="TabbedPanels" id="tp1" align='center'>
				
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">RKB WIP</li>
						
					</ul>

					<div class="TabbedPanelsContentGroup">

						<!-- view bkm wip-->
						<div class="TabbedPanelsContent">
							<table align=center border="1" width="900" class='data' id='table2'>
								<tr id="trheader2">
									<th>Aktivitas</th>
									<th>Qty Kontrak</th>
									<th>Harga</th>
									<th>Blok</th>
									<th>RKB Qty</th>
									<th>Total</th>
								
								</tr>
								<tr id='tru'>
									<tbody id="bodytablewip">
									</tbody>
								</tr>

								<tr class="ratatengah">
										<td colspan="4" class="tebal">Total</td>
										<td class="tebal">
											<input class="ratatengah lebarupah" type="text" readonly  style="background-color:#eee;" name="amountrkbqty" id="amountrkbqtyfrm" value="0"/>
											<input class="ratatengah lebarupah" type="hidden" readonly  style="background-color:#eee;" name="amountrkbqty" id="amountrkbqty" value="0"/>
										</td>
										<td class="tebal">
											<input class="ratakanan lebarallow" type="text" readonly  style="background-color:#eee;" name="amountrkbrpfrm" id="amountrkbrpfrm" value="0"/>
											<input class="ratakanan lebarallow" type="hidden" readonly  style="background-color:#eee;" name="amountrkbrp" id="amountrkbrp" value="0"/>
										</td>
								</tr>
								<tr>
									<td  colspan="6" class="ratatengah"><input type=Submit name="simpan" value="SAVE"></td>
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
</body>
<!--#include file="bottom.jsp"-->
</html>
<link href="/js/select2/select2.min.css" rel="stylesheet" type="text/css"/>
<script src="/js/select2/select2.min.js" type="text/javascript"></script>

<script type="text/javascript">
	new Spry.Widget.TabbedPanels("tp1", {
		defaultTab: 0
	});

	function hanyaAngka(evt) {
		evt.value = evt.value.replace(/[^0-9,.-]/g,"");
	}

$(function() {	
	
	showKontrak();
	
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

//get  agreemnt code
function getAgreementcode() {
	var fcbaselect = $("#fcbawip").val();
	$('#nokontrak').empty();
	$('#nokontrak').select2({
		ajax: {
			delay: 1000,
			url : "bkm_show_item.asp?contractorcode=" + contractorcode + "&jmlitem=" + 11 + "&fcbaselect=" + fcbaselect,
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

function getFCBAar()
{
	var company = $("#companyar").val();
	showFCBAar(company);
	showFCBAWIP(company);

}

function showFCBAar(company) {
			$('#businessunitar').empty();
			$('#businessunitar').select2({
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

//TAMPLKAN FCBA DARI INPUTAN RKB WIP
function showFCBAWIP(company) {
			$('#chargeunit').empty();
			$('#chargeunit').select2({
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

function showKontrak(businessunitar) {
			$('#nokontrak').empty();
			$('#nokontrak').select2({
				ajax: {
					delay: 1000,
					url : 'bkm_show_item.asp?fcbaselect=' + businessunitar + '&jmlitem=' + 13 + '&tahun=' + $('#tahun').val() + '&bulan=' + $('#bulan').val(),
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

function getKontrak()
{
	var businessunitar = $("#businessunitar").val();
	showKontrak(businessunitar);
}

 //ambil data detail
function viewData()
	{
		var companyar 		= $("#companyar").val();
		var businessunitar 	= $("#businessunitar").val();
		var nokontrak 		= $("#nokontrak").val();
		if (nokontrak != "") { 
				$('#loadingjudul').show();
				$('#loading').show();
				$.post("bkm_viewdata.asp?company=" + companyar + "&fcbarkb=" + businessunitar + "&nokontrak=" + nokontrak + "&idview=" + 29, function(data) {
				$('#bodytablewip').empty();
				$('#contractorcode').empty();
				$('#contractorname').empty();
				var json  = JSON.parse(data);
				var contractorcode = json.results[0].contractorcode;
				var contractorname = json.results[0].contractorname;
				$('#contractorcode').val(contractorcode);
				$('#contractorname').val(contractorname);
				var jumlview  = json.results.length;
				var no = 1;
				for (var i = 0; i < jumlview; i++) {
					var docketlist = '';
						docketlist += '<tr id="trbs'+i+'">';
						docketlist += '<td align="left">' + json.results[i].jobcode + ' - ' + json.results[i].fcname + '<input class="ratakanan lebarupah" type="hidden" name="actcode' + i + '" id="actcode' + i + '" value="' + json.results[i].jobcode + '"/></td>';
						docketlist += '<td align="right">' + json.results[i].qty + '</td>';
						docketlist += '<td align="right">' + numformat(json.results[i].rate) + '<input class="ratakanan lebarupah" type="hidden" name="rate' + i + '" id="rate' + i + '" value="' + json.results[i].rate + '"/></td>';
						docketlist += '<td align="center"><select class="lebarupah" name="blok' + i + '" id="blok' + i + '" value=""/></select></td>';
						docketlist += '<td align="center"><input class="ratatengah lebarupah" onkeyup="hanyaAngka(this);getSubtotal(' + i + ');getPuttotal();getPuttotalrp()" type="text" name="rkbqty' + i + '" id="rkbqty' + i + '" value="0"/></td>';
						docketlist += '<td align="center"><input class="ratakanan lebarallow" type="text" name="subtotalfrm' + i + '" id="subtotalfrm' + i + '" value="0"/><input class="ratakanan lebarallow" type="hidden" name="subtotal' + i + '" id="subtotal' + i + '" value="0"/><input class="ratakanan lebarallow" type="hidden" name="lineno' + i + '" id="lineno' + i + '" value="' + json.results[i].lineno + '"/></td>';
						docketlist += '</tr>';
						$('#bodytablewip').append(docketlist);
					no++;
					$("#jmlrkb").val(jumlview);
					getBlok(i);
					//new Spry.Widget.ValidationSelect("blok" + i);
					new Spry.Widget.ValidationTextField("rkbqty" + i);
				}

			});
		 
		} else {
			$('#bodytablewip').empty();
		}
	}

	//get tc
	function getBlok(i)
	{
		$('#blok' + i).empty();
		$('#blok' + i).select2({
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

	function getSubtotal(i)
	{
		var harga 		= parseInt($('#rate' + i).val());
		var rkbqty 		= parseInt($('#rkbqty' + i).val());
		//alert(harga);
		var total 		= harga*rkbqty;
		var totalfrm    = numformat(total);
		$('#subtotalfrm' + i).val((totalfrm));
		$('#subtotal' + i).val((total));
	}
	
	function getPuttotal() {
		var ttlrkb	= 0;
		var jmlrkb	= parseInt($('#jmlrkb').val());
		for(var i=0; i<jmlrkb; i++) {
			ttlrkb	    += parseFloat($('#rkbqty'+i).val());
		  
			
		}	
		ttlrkbfrm     = numformat(ttlrkb);
		$('#amountrkbqtyfrm').val((ttlrkbfrm));
		$('#amountrkbqty').val((ttlrkb));
		
	}

	function getPuttotalrp() {
		var subtotal	= 0;
		var jmlrkb	= parseInt($('#jmlrkb').val());
		for(var i=0; i<jmlrkb; i++) {
			subtotal	    += parseFloat($('#subtotal'+i).val());
		}	
		var subtotalfrm     = numformat(subtotal);
		$('#amountrkbrpfrm').val((subtotalfrm));
		$('#amountrkbrp').val((subtotal));
		
	}

</script>
