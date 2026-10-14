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

'count input
Dim SQL_Rows, NoOfRows,RSCount

SQL_Rows = "SELECT Count (*) from TANAMANMST "
set RSCount = connStr.Execute(SQL_Rows)
if not RSCount.eof then 
	NoOfRows = RSCount(0).value
end if
RSCount.close



%>

<html>
	<head>
		<title>FBPlants - HARIAN TANAMAN</title>
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

<body>
<!--#include file="top.jsp"-->

<br>
<br>
<br>
<br>
<br>


<div id="head_judul" style="float:left;font-size:10px; text-align:left;">
		<strong class="posisikiri">&nbsp;&nbsp;PT. <%=CompanyName%> - &nbsp;<%=FcbaName%>   -           MODULE : HARIAN TANAMAN</strong><br>					
</div>
<br>
<br>
<div class="box_bothall" style="overflow:scroll;">			
<form name="frmEntry" method="post" action="bkm_tanaman_proses.asp">
	
	<input type="hidden" name="jumldata" id="jumldata" value="<%=NoOfRows%>" />
	<table align=center class="data ratatengah" border="0" width="60%">			
		<tr>
			<td align="left" >&nbsp;&nbsp;&nbsp;Tanggal</td>
			<td align="left">
				<input type="text" class="datepicker" maxlength="15" name="tglinput" readonly id="tglinput" size="15" value="<%=textDate(date)%>" title="Tanggal">
			</td>			
		</tr>
	</table>
	
	<br>
		
	<table align=center border="0" class='data posisikiri' width='700' id='tableall'>
		<tr>
			<td colspan=>
				<div class="TabbedPanels" id="tp1" align='center'>
					<ul class="TabbedPanelsTabGroup">
						<li class="TabbedPanelsTab">INPUT HARIAN</li>
						
					</ul>

					<div class="TabbedPanelsContentGroup">

						<div class="TabbedPanelsContent">
							<table align=center border="0" width="700" class='data' id='table2'>
								<tr id="trheader2">
									
									<th>Group</th>
									<th>Uraian</th>
									<th>Satuan</th>
									<th>Jumlah</th>
									<th colspan="2">&nbsp;</th>
								</tr>
								<% 
								
								
								'get data detail upah
								dim sqldtl,rsdtl,no,totalamount
								sqldtl = "select g.uraian,t.kode,t.kode || ' - ' || t.uraian item,t.satuan " & _
								" from TANAMANGROUP g " & _
								" inner join TANAMANMST t on (g.id=t.groupmst) " & _
								" order by t.groupmst,t.kode" 
								set rsdtl = connStr.Execute(sqldtl)
								'response.write sqldtl
								
																
								no = 1
								totalamount = 0
								While Not rsdtl.EOF %>

								<tr id="tru<%=no%>">
									
									<td class="ratakiri"><%=rsdtl("uraian")%>
										<input class="ratakiri lebaremployee" readonly type="hidden" name="group<%=no%>" id="group<%=no%>" value="<%=rsdtl("uraian")%>"/>
									</td>
									<td class="ratakiri"><%=rsdtl("item")%>
										<input class="ratakiri lebaremployee" readonly type="hidden" name="kodeitem<%=no%>" id="kodeitem<%=no%>" value="<%=rsdtl("kode")%>"/>
									</td>
									<td class="ratatengah"><%=rsdtl("satuan")%>
										<input class="ratatengah lebarrate" readonly type="hidden" name="satuan<%=no%>" id="satuan<%=no%>" value="<%=rsdtl("satuan")%>"/>
									</td>
									<td class="ratatengah">
										<input class="ratatengah lebarrate" type="text" name="jumlah<%=no%>" id="jumlah<%=no%>" value="0"/>
									</td>
									
									<td></td>
									
								</tr>

								
								<%

									

									no = no +1
									'totalamount = cdbl(totalamount) + cdbl(rsdtlupah("amountdebit"))
									rsdtl.MoveNext
									
									Wend
								%>
								
							</table>
						</div>
					</div>
				</div>
			</td>
		</tr>
		<tr>
			<td colspan='' align=center><input Type="Submit"  name="simpan" value="Simpan"></td>
			
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
	$("#tglinput").datepicker();
	
	
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

//new Spry.Widget.ValidationTextField("DATECREATED", "none");
//new Spry.Widget.ValidationSelect("clasification");
</script>