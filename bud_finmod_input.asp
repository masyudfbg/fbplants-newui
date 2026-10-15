<%@Language=VBScript%>
<% Option Explicit %>
<% ValidateSession %>
<!--#include virtual="/connDb1.Inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->

<%
dim businessunit,fcbaname,CompanyName,q,sqltotdata,gangcode,pcompanycode,ptahun
pcompanycode= trim(request("company"))

if trim(request("year")) = "" then 
ptahun      = YEAR(textDate(Date))
else
ptahun	    = trim(request("year"))
end if
%>

<html>
	<head>
		<title>FBPlants - FINANCIAL MODEL</title>
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
<script type="text/javascript" src="/js/fbp-shell.js?v=20261015a"></script>
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
		<strong class="posisikiri">MODULE : BUDGET - FINANCIAL MODEL</strong><br>					
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
			<select name="tahun" id="tahun" style="width:120px">
			<%  
			    Dim sqltahun5,rstahun5
				
				sqltahun5 = "select distinct to_number(accountingyear) tahun from accountingyear order by tahun "  
				Set rstahun5 = Server.CreateObject("ADODB.Recordset")
				rstahun5.CursorType = 3
				rstahun5.CursorLocation = 3
				rstahun5.Open sqltahun5, connstr
				While Not rstahun5.EOF 
				    if cint(ptahun)=cint(rstahun5("tahun")) then %>
					<option value="<%=rstahun5("tahun")%>" Selected ><%=rstahun5("tahun")%></option>
					<%else%>
					<option value="<%=rstahun5("tahun")%>"><%=rstahun5("tahun")%></option>
					<%end if%>					
				<%rstahun5.movenext
				wend 
				rstahun5.close 
				

			%>	
			</select>
			</td>
		</tr>
		<tr>
			<td align="left" >Company</td>
			<td align="left">
				<select name="company" id="company" style="width:120px" onchange="getEmployee();" >
			            <option value=""></option>
						<% 	Dim rsCompany
							q = " SELECT companycode, companyname " & _
							 	" FROM company a INNER JOIN user_company_tab b ON a.companycode=b.fccompanycode " & _
								" WHERE b.fcusercode='" & Session("CodeName") & "'" & _								
							    " order by companyname"
							 response.write q 
							Set rsCompany = Server.CreateObject("ADODB.Recordset")
							rsCompany.CursorType = 3
							rsCompany.CursorLocation = 3
							rsCompany.Open q, connstr
							if not rsCompany.eof then%>
							<%
							Do While Not rsCompany.eof
								if trim(pcompanycode)=trim(rsCompany("companycode")) then %>
							%>
								<option value="<%=rsCompany("companycode")%>" Selected ><%=rsCompany("companycode")%></option> 
							<% else %>
								<option value="<%=rsCompany("companycode")%>"><%=rsCompany("companycode")%></option> 
							<%end if 
							rsCompany.movenext
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
				<div class="TabbedPanels" id="tp1" align='center'>
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">GENERAL</li>
						<li class="TabbedPanelsTab">BIBITAN</li>
						<li class="TabbedPanelsTab">LC</li>
						<li class="TabbedPanelsTab">TBM</li>
						<li class="TabbedPanelsTab">TM</li>
						<li class="TabbedPanelsTab">POM</li>
						
					</ul>

					<div class="TabbedPanelsContentGroup">
						<div class="TabbedPanelsContent">
						<!-- GENERAL ============================================================================================================================= -->
							<table align=center border="0" width="1200" class='data' id='table2'>
								<tr id="trheader2">
									
									<th>Type</th>
									<th>Ownership</th>
									<th>Code</th>
									<th>Name</th>
									<th>Ha.Gross</th>
									<th>Qty.FinMod</th>
									<th>Price</th>
									<th>Amount</th>
									<th>Remark</th>
									
								</tr>
								<tr id='tru'>
									<tbody id="bodytable_GE">
									</tbody>
								</tr>
		
							</table>
						</div>
						
						<div class="TabbedPanelsContent">
						<!--  BIBITAN ============================================================================================================================= -->
							<table align=center border="0" width="1200" class='data' id='table2'>
								<tr id="trheader2">
									
									<th>Type</th>
									<th>Ownership</th>
									<th>Code</th>
									<th>Name</th>
									<th>Ha.Gross</th>
									<th>Qty.FinMod</th>
									<th>Price</th>
									<th>Amount</th>
									<th>Remark</th>
									
								</tr>
								<tr id='tru'>
									<tbody id="bodytable_NU">
									</tbody>
								</tr>
		
							</table>
						</div>

						<div class="TabbedPanelsContent">
						<!-- LC ============================================================================================================================= -->
							<table align=center border="0" width="1200" class='data' id='table2'>
								<tr id="trheader2">
									
									<th>Type</th>
									<th>Ownership</th>
									<th>Code</th>
									<th>Name</th>
									<th>Ha.Gross</th>
									<th>Qty.FinMod</th>
									<th>Price</th>
									<th>Amount</th>
									<th>Remark</th>
									
								</tr>
								<tr id='tru'>
									<tbody id="bodytable_LC">
									</tbody>
								</tr>

								
							</table>
						</div>
						
						<div class="TabbedPanelsContent">
						<!-- TBM ============================================================================================================================= -->
							<table align=center border="0" width="1200" class='data' id='table2'>
								<tr id="trheader2">
									
									<th>Type</th>
									<th>Ownership</th>
									<th>Code</th>
									<th>Name</th>
									<th>Ha.Gross</th>
									<th>Qty.FinMod</th>
									<th>Price</th>
									<th>Amount</th>
									<th>Remark</th>
									
								</tr>
								<tr id='tru'>
									<tbody id="bodytable_TBM">
									</tbody>
								</tr>

							
								
							</table>
						</div>
						
						<div class="TabbedPanelsContent">
						<!-- TM ============================================================================================================================= -->
							<table align=center border="0" width="1200" class='data' id='table2'>
								<tr id="trheader2">
									
									<th>Type</th>
									<th>Ownership</th>
									<th>Code</th>
									<th>Name</th>
									<th>Ha.Gross</th>
									<th>Qty.FinMod</th>
									<th>Price</th>
									<th>Amount</th>
									<th>Remark</th>
									
								</tr>
								<tr id='tru'>
									<tbody id="bodytable_TM">
									</tbody>
								</tr>

								
								
							</table>
						</div>
						
						<div class="TabbedPanelsContent">
						<!-- POM ============================================================================================================================= -->
							<table align=center border="0" width="1200" class='data' id='table2'>
								<tr id="trheader2">
									
									<th>Type</th>
									<th>Ownership</th>
									<th>Code</th>
									<th>Name</th>
									<th>Ha.Gross</th>
									<th>Qty.FinMod</th>
									<th>Price</th>
									<th>Amount</th>
									<th>Remark</th>
									
								</tr>
								<tr id='tru'>
									<tbody id="bodytable_POM">
									</tbody>
								</tr>
		
							</table>
						</div>
					
							<!-- ============================================================================================================================= -->
					</div>
				</div>
			</td>
		</tr>
		<tr>
									<td colspan='15' align=center><input Type="Submit" id="SIMPAN_ALL" name="SIMPAN" value="SIMPAN"  /> </td>
									
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
	//$("#tglbkm").datepicker();
	//showSection();
	//showClasification();
	//showSupervisi1();
	//showSupervisi2();
	//showSupervisi3();
	//showSupervisi4();
	//addrowBahan(1);

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
function getEmployee()
{
	var pcompany = $("#company").val();
	var ptahun = $('#tahun').val();
	if (pcompany != "") {
		$.post("bud_finmod_getjobgroup.asp?company=" + pcompany + "&year=" + ptahun + "&id=" + 2, function(data) {
			$('#bodytable_GE').empty();
			$('#bodytable_NU').empty();
			$('#bodytable_LC').empty();
			$('#bodytable_TBM').empty();
			$('#bodytable_TM').empty();
			$('#bodytable_POM').empty();
			
			var json  = JSON.parse(data);
			var jumlview  = json.results.length;
			var no = 1;
			if(jumlview>0){
			for (var i = 0; i < jumlview; i++) {
			 	var job_forfieldstatus = json.results[i].job_forfieldstatus;
			    var ownership = json.results[i].ownership;
				var fccode = json.results[i].jobgroupcode;
				var fcname = json.results[i].jobgroupname;
				var qtyunit = json.results[i].qtyunit;
				
				var qty = json.results[i].qty;
				var price = json.results[i].price;
				var amount = json.results[i].amount;
				var remarks = json.results[i].remarks;
				
				
				
				var docketlist = '';
					docketlist += '<tr id="tru'+no+'">';
					docketlist += '<td align="Left">' + job_forfieldstatus + '<input type="hidden" class="ratakiri lebarhk" name="job_forfieldstatus' + no + '" id="job_forfieldstatus' + no + '" value="' + job_forfieldstatus + '"/></td>';
					docketlist += '<td align="Left">' + ownership + '<input type="hidden" class="ratakiri lebarhk" name="ownership' + no + '" id="ownership' + no + '" value="' + ownership + '"/></td>';
					docketlist += '<td align="Left">' + fccode + '<input type="hidden" class="ratakiri lebarhk" name="fccode' + no + '" id="fccode' + no + '" value="' + fccode + '"/></td>';
					docketlist += '<td align="Left">' + fcname + '<input type="hidden" class="ratakiri lebarhk" name="fcname' + no + '" id="fcname' + no + '" value="' + fcname + '"/></td>';
					docketlist += '<td align="right">' + numformat(qtyunit) + '<input type="hidden" class="ratakanan lebarhk" name="qtyunit' + no + '" id="qtyunit' + no + '" value="' + qtyunit + '"/></td>';
					
					docketlist += '<td align="right"><input cols="10" class="ratakanan  type="text" onkeyup="hanyaAngka(this)" name="qty' + no + '" id="qty`' + no + '" value="' + qty + '"/></td>';
					docketlist += '<td align="right"><input cols="10" class="ratakanan  type="text" onkeyup="hanyaAngka(this)" name="price' + no + '" id="price' + no + '" value="' + price+ '"/></td>';
					docketlist += '<td align="right"><input cols="10" class="ratakanan  type="text" disabled name="amount' + no + '" id="amount' + no + '" value="' + amount+ '"/></td>';
					docketlist += '<td align="Left"><input cols="10" type="text"  name="remark' + no + '" id="remark' + no + '" value="' + remarks + '"/></td>';
					
					docketlist += '</tr>';
					if (job_forfieldstatus == "GENERAL") {
					$('#bodytable_GE').append(docketlist);
					}
					
					
					if (job_forfieldstatus == "BIBITAN") {
					$('#bodytable_NU').append(docketlist);
					}
					
					if (job_forfieldstatus == "LC") {
					$('#bodytable_LC').append(docketlist);
					}
					
					if (job_forfieldstatus == "TBM") {
					$('#bodytable_TBM').append(docketlist);
					}
					
					if (job_forfieldstatus == "TM") {
					$('#bodytable_TM').append(docketlist);
					}
				
					if (job_forfieldstatus == "POM") {
					$('#bodytable_POM').append(docketlist);
					}
				no++;
				$("#jumldata").val(jumlview);
			}
			
			}else{
				getEmployeeview();
			}
			
			
						
		});
		
	} 
}


//getEmployee
function getEmployeeview() {
	var pcompany = $("#company").val();
	
	if (pcompany != "") {
		$.post("bud_finmod_getjobgroup.asp?company="+ pcompany + "&id=" + 1, function(data) {
		$('#bodytable_GE').empty();
		$('#bodytable_NU').empty();
		$('#bodytable_LC').empty();
		$('#bodytable_TBM').empty();
		$('#bodytable_TM').empty();
		$('#bodytable_POM').empty();
		
			var json  = JSON.parse(data);
			var jumlview  = json.results.length;
			var no = 1;
			for (var i = 0; i < jumlview; i++) {
				var job_forfieldstatus = json.results[i].job_forfieldstatus;
				var ownership = json.results[i].ownership;
				var fccode = json.results[i].fccode;
				var fcname = json.results[i].fcname;
				var qtyunit = json.results[i].qtyunit;
				
				
				var docketlist = '';
					docketlist += '<tr id="tru'+no+'">';
					docketlist += '<td align="left">' + job_forfieldstatus + '<input type="hidden" class="ratakiri lebarhk" name="job_forfieldstatus' + no + '" id="job_forfieldstatus' + no + '" value="' + job_forfieldstatus + '"/></td>';
					docketlist += '<td align="left">' + ownership + '<input type="hidden" class="ratakiri lebarhk" name="ownership' + no + '" id="ownership' + no + '" value="' + ownership + '"/></td>';
					docketlist += '<td align="left">' + fccode + '<input type="hidden" class="ratakiri lebarhk" name="fccode' + no + '" id="fccode' + no + '" value="' + fccode + '"/></td>';
					docketlist += '<td align="left">' + fcname + '<input type="hidden" class="ratakiri lebarhk" name="fcname' + no + '" id="fcname' + no + '" value="' + fcname + '"/></td>';
					docketlist += '<td align="right">' + numformat(qtyunit) + '<input type="hidden" class="ratakanan lebarhk" name="qtyunit' + no + '" id="qtyunit' + no + '" value="' + qtyunit + '"/></td>';
					
					docketlist += '<td align="right"><input cols="10" type="text" class="ratakanan  onkeyup="hanyaAngka(this)" name="qty' + no + '" id="qty`' + no + '" value="0"/></td>';
					docketlist += '<td align="right"><input cols="10" type="text" class="ratakanan  onkeyup="hanyaAngka(this)" name="price' + no + '" id="price' + no + '" value="0"/></td>';
					docketlist += '<td align="right"><input cols="10" type="text" class="ratakanan  disabled name="amount' + no + '" id="amount' + no + '" value="0"/></td>';
					docketlist += '<td align="left"><input cols="10" type="text"  name="remark' + no + '" id="remark' + no + '" value=""/></td>';
					
					docketlist += '</tr>';
					if (job_forfieldstatus == "GENERAL") {
					$('#bodytable_GE').append(docketlist);
					}
					
					if (job_forfieldstatus == "BIBITAN") {
					$('#bodytable_NU').append(docketlist);
					}
					
					if (job_forfieldstatus == "LC") {
					$('#bodytable_LC').append(docketlist);
					}
					
					if (job_forfieldstatus == "TBM") {
					$('#bodytable_TBM').append(docketlist);
					}
					
					if (job_forfieldstatus == "TM") {
					$('#bodytable_TM').append(docketlist);
					}
				
					if (job_forfieldstatus == "POM") {
					$('#bodytable_POM').append(docketlist);
					}
					
				no++;
				$("#jumldata").val(jumlview);
			}
			
		});
		
	} else {
		$('#bodytable_GE').empty();
		$('#bodytable_NU').empty();
		$('#bodytable_LC').empty();
		$('#bodytable_TBM').empty();
		$('#bodytable_TM').empty();
		$('#bodytable_POM').empty();
		
	}
	
}


</script>