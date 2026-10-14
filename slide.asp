
<html>
    <head>
    <style>

/* table.tableclass {
  border-collapse: collapse;
  font-size: 80%;
  font-family: "trebuchet ms", "lucidia grande", verdana, arial, sans-serif;

}

table.tableclass thead th {
	border-top: 1px solid #D4E0EE;
	text-align: center;
	background: #f4f4f4;
	color: #414476;
}

table.tableclass tbody tr td, table.tableclass thead tr th {
  border-bottom: 1px solid black;
  padding: 0.1rem;
}
.bgtbs {
    background:#B6FCA8;
}
.bgcpo {
    background:#FCE3A8;
}
.bgkernel {
    background:#FCA8F7;
}
.bgcangkang {
    background:#A8C6FC;
}
.bgtbsjual {
    background:#96D030;
}
.bgampas {
    background:#F8B2B3;
}
.bgabu {
    background:#F8F8B2;
}
.txtimbanganlain {
    color:red;
} */


.table1 {
    font-family: sans-serif;
    color: #444;
    border-collapse: collapse;
    width: 90%;
    border: 1px solid #f2f5f7;
	font-size: 13px;
	
}
 
.table1 tr th{
    background: #35A9DB;
    color: #fff;
    font-weight: normal;
}
 
.table1, th {
    padding: 4px 15px;
    text-align: center;
}
.table1, td {
    padding: 4px 15px;
   
}
 
.table1 tr:hover {
    background-color: #f5f5f5;
}
 
.table1 tr:nth-child(even) {
    background-color: #f2f2f2;
}



.table2 {
    font-family: sans-serif;
    color: #444;
    border-collapse: collapse;
    width: 80%;
    border: 1px solid #f2f5f7;
}
 
.table2 tr th{
    background: #35A9DB;
    color: #fff;
    font-weight: normal;
}
 
.table2, th {
    padding: 4px 15px;
    text-align: center;
}
.table2, td {
    padding: 4px 15px;
   
}
 
.table2 tr:hover {
    background-color: #f5f5f5;
}
 
.table2 tr:nth-child(even) {
    background-color: #f2f2f2;
}
.lebarket{
	width:190px;
}
</style>
    </head>
	
	<script src="/js/jquery-1.9.1.js"></script>
	
<script>
	$(function() {
		getSummaryTBS();
		getInkendaraan();
	});

	var timerKendaraanVar = setInterval(function() { getSummaryTBS() }, 120000);
	//var timerKendaraanVar = setInterval(function() { getInkendaraan() }, 30000);

	//GET QTY TBS
	function getSummaryTBS()
	{			$('#bodytable').empty();
				//get TBS
				$.post("getporeqpayview.asp?validasi=" + 3, function(data) {
					
					var json  = JSON.parse(data);
					var jumlviewtbs  = json.dataviewtbs.length;
					
					for (var i = 0; i < jumlviewtbs; i++) {
						var areatypename 	= json.dataviewtbs[i].areatypename;
						var tonasearmhi 	= json.dataviewtbs[i].tonasearmhi;
						var tonasearmshi 	= json.dataviewtbs[i].tonasearmshi;
						var tonasearmsbi 	= json.dataviewtbs[i].tonasearmsbi;
						var tonaseramhi		= json.dataviewtbs[i].tonaseramhi
						var tonaseramshi 	= json.dataviewtbs[i].tonaseramshi;
						var tonaseramsbi 	= json.dataviewtbs[i].tonaseramsbi;
						var tonasebrehi 		= json.dataviewtbs[i].tonasebrehi;
						var tonasebreshi 		= json.dataviewtbs[i].tonasebreshi;
						var tonasebresbi 		= json.dataviewtbs[i].tonasebresbi;
						if (jumlviewtbs>0){
							var docketlist = '';
							
							docketlist += '<tr>';
							docketlist += '<td align="left" class="lebarket">TBS Masuk - ' + json.dataviewtbs[i].areatypename + '</td>';
							docketlist += '<td align="right">' + numformat(tonasearmhi) + '</td>';
							docketlist += '<td align="right">' + numformat(tonasearmshi) + '</td>';
							docketlist += '<td align="right">' + numformat(tonasearmsbi) + '</td>';
							docketlist += '<td align="right">' + numformat(tonaseramhi) + '</td>';
							docketlist += '<td align="right">' + numformat(tonaseramshi) + '</td>';
							docketlist += '<td align="right">' + numformat(tonaseramsbi) + '</td>';
							docketlist += '<td align="right">' + numformat(tonasebrehi) + '</td>';
							docketlist += '<td align="right">' + numformat(tonasebreshi) + '</td>';
							docketlist += '<td align="right">' + numformat(tonasebresbi) + '</td>';
							docketlist += '</tr>';
							$('#bodytable').append(docketlist);
						}
						
					}
				});

				
				//get CPO
				$.post("getporeqpayview.asp?validasi=" + 4, function(data) {
					var json  = JSON.parse(data);
					var jumlviewcpo  = json.dataviewcpo.length;
					for (var x = 0; x < jumlviewcpo; x++) {
						
						var totalberatcpo 		= (json.dataviewcpo[x].tonasecpoarmhi);
						var tonasecpoarmshi 		= (json.dataviewcpo[x].tonasecpoarmshi);
						var tonasecpoarmsbi 		= (json.dataviewcpo[x].tonasecpoarmsbi);
						//if (totalberatcpo>0){
							
							var docketlist = '';
							docketlist += '<tr>';
							docketlist += '<td align="left" class="lebarket">CPO Keluar</td>';
							docketlist += '<td align="right">' + numformat(totalberatcpo) + '</td>';
							docketlist += '<td align="right">' + numformat(tonasecpoarmshi) + '</td>';
							docketlist += '<td align="right">' + numformat(tonasecpoarmsbi) + '</td>';
							docketlist += '<td align="right">' + '0' + '</td>';
							docketlist += '<td align="right">' + '0' + '</td>';
							docketlist += '<td align="right">' + '0' + '</td>';
							docketlist += '<td align="right">' + '0' + '</td>';
							docketlist += '<td align="right">' + '0' + '</td>';
							docketlist += '<td align="right">' + '0' + '</td>';
							docketlist += '</tr>';
							$('#bodytable').append(docketlist);
						//}
						
					}
				});

				//get PK
				$.post("getporeqpayview.asp?validasi=" + 5, function(data) {
					var json  = JSON.parse(data);
					var jumlviewpk  = json.dataviewpk.length;
					for (var p = 0; p < jumlviewpk; p++) {
						
						var tonasepkarmhi 		= (json.dataviewpk[p].tonasepkarmhi);
						var tonasepkarmshi 		= (json.dataviewpk[p].tonasepkarmshi);
						var tonasepkarmsbi 		= (json.dataviewpk[p].tonasepkarmsbi);
						//if (totalberatpk>0){
							
							var docketlist = '';
							docketlist += '<tr>';
							docketlist += '<td align="left" class="lebarket">Kernel Keluar</td>';
							docketlist += '<td align="right">' + numformat(tonasepkarmhi) + '</td>';
							docketlist += '<td align="right">' + numformat(tonasepkarmshi) + '</td>';
							docketlist += '<td align="right">' + numformat(tonasepkarmsbi) + '</td>';
							docketlist += '<td align="right">' + '0' + '</td>';
							docketlist += '<td align="right">' + '0' + '</td>';
							docketlist += '<td align="right">' + '0' + '</td>';
							docketlist += '<td align="right">' + '0' + '</td>';
							docketlist += '<td align="right">' + '0' + '</td>';
							docketlist += '<td align="right">' + '0' + '</td>';
							docketlist += '</tr>';
							$('#bodytable').append(docketlist);
						//}
						
					}
				});

				//get CKG
				$.post("getporeqpayview.asp?validasi=" + 6, function(data) {
					var json  = JSON.parse(data);
					var jumlviewckg  = json.dataviewckg.length;
					for (var p = 0; p < jumlviewckg; p++) {
						
						var tonaseckgarmhi 		= (json.dataviewckg[p].tonaseckgarmhi);
						var tonaseckgarmshi 	= (json.dataviewckg[p].tonaseckgarmshi);
						var tonaseckgarmsbi 	= (json.dataviewckg[p].tonaseckgarmsbi);
					//	if (totalberatckg>0){
							
							var docketlist = '';
							docketlist += '<tr>';
							docketlist += '<td align="left" class="lebarket">Cangkang Keluar</td>';
							docketlist += '<td align="right">' + numformat(tonaseckgarmhi) + '</td>';
							docketlist += '<td align="right">' + numformat(tonaseckgarmshi) + '</td>';
							docketlist += '<td align="right">' + numformat(tonaseckgarmsbi) + '</td>';
							docketlist += '<td align="right">' + '0' + '</td>';
							docketlist += '<td align="right">' + '0' + '</td>';
							docketlist += '<td align="right">' + '0' + '</td>';
							docketlist += '<td align="right">' + '0' + '</td>';
							docketlist += '<td align="right">' + '0' + '</td>';
							docketlist += '<td align="right">' + '0' + '</td>';
							docketlist += '</tr>';
							$('#bodytable').append(docketlist);
					//	}
						
					}
				});

				//get RETAIL
				// $.post("getporeqpayview.asp?validasi=" + 7, function(data) {
					// var json  = JSON.parse(data);
					// var jumlviewretail  = json.dataviewretail.length;
					// for (var r = 0; r < jumlviewretail; r++) {
						
						// var tonaselainarmhi 	= (json.dataviewretail[r].tonaselainarmhi);
						// var tonaselainarmshi 	= (json.dataviewretail[r].tonaselainarmshi);
						// var tonaselainarmsbi 	= (json.dataviewretail[r].tonaselainarmsbi);

						// if (totalberatretail>0){
							// totalberatretail		= numformat(totalberatretail);
							// var docketlist = '';
							// docketlist += '<tr>';
							// docketlist += '<td align="left" class="lebarket">Retail/Lain-Lain Keluar</td>';
							// docketlist += '<td align="right">' + numformat(tonaselainarmhi) + '</td>';
							// docketlist += '<td align="right">' + numformat(tonaselainarmshi) + '</td>';
							// docketlist += '<td align="right">' + numformat(tonaselainarmsbi) + '</td>';
							// docketlist += '<td align="right">' + '0' + '</td>';
							// docketlist += '<td align="right">' + '0' + '</td>';
							// docketlist += '<td align="right">' + '0' + '</td>';
							// docketlist += '<td align="right">' + '0' + '</td>';
							// docketlist += '<td align="right">' + '0' + '</td>';
							// docketlist += '<td align="right">' + '0' + '</td>';
							// docketlist += '</tr>';
							// $('#bodytable').append(docketlist);
						// }
						
					// }
				// });
				

		
			
	}

	

	function getInkendaraan()
	{
				var validasi =8;
				$.post("getporeqpayview.asp?validasi=" + validasi, function(data) {
					$('#bodytable1').empty();
					var json  = JSON.parse(data);
					var jumlview  = json.datain.length;
					for (var i = 0; i < jumlview; i++) {

						var fcba 	= json.datain[i].fcba;
						var DocketID 	= json.datain[i].docketid;
						var NoReg 		= json.datain[i].noreg;
						var DriverName 	= json.datain[i].drivername;
						var TimeIn 		= json.datain[i].timein;
						var WeightIn 	= (json.datain[i].weightin);
						var InWBID 		= json.datain[i].inwbid;
						var productid 	= json.datain[i].productid;
						var ProductName	= json.datain[i].productname;

						var prod;

						if (productid == '1') {
								prod		= "CPO";
							} else if (productid == '2') {
								prod		= "TBS";
							} else if (productid == '3') {
								prod		= "KERNEL";
							} else if ($productid == '4') {
								prod		= "AMPAS";
							} else if (productid == '5') {
								prod		= "CANGKANG";
							} else if (productid == '6') {
								prod		= "KOMPOS";
							} else if (productid == '7') {
								prod		= "TBS JUAL";
							} else if (productid == '10') {
								prod		= ProductName;
							} else {
								prod		= "";
							}

							//alert(WeightIn);

						var docketlist = '';
						 docketlist += '<tr>';
						 docketlist += '<td align="left">' + fcba + '</td>';
						 docketlist += '<td align="left">' + prod + '</td>';
						 docketlist += '<td align="left">' + DocketID + '</td>';
						 docketlist += '<td align="left">' + NoReg + '</td>';
						 docketlist += '<td align="left">' + DriverName + '</td>';
						 docketlist += '<td align="left">' + TimeIn + '</td>';
						 docketlist += '<td align="right">' + WeightIn + '</td>';
						 docketlist += '</tr>';
						 $('#bodytable1').append(docketlist);
					}
				});
				
			
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
</script>
<body>
<%
'query monitoring timbang TBS'
dim tglwbnow,sqlwb,rswb,fcbasession,wbdate
fcbasession = Session("fcBA")
sqlwb = "select WBDATE FROM configmst"
Set rswb = Server.CreateObject("ADODB.Recordset")
rswb.CursorType = 3
rswb.CursorLocation = 3
rswb.Open sqlwb, connstr
if not rswb.eof then
    wbdate = rswb("WBDATE")
    
end if
%>
<br>
<div align="center">
	<span style="font-size:22px;"><strong>MONITORING JEMBATAN TIMBANG</strong></span>
</div>
<br>
<div align="center">
	<span style="font-size:14px;"><strong>Summary Timbang Hari Ini: <%=wbdate%></strong></span>
</div>
<br>
<div align="center">
<table class="table1" border=1  width="85%" >
	 <tr>
	 	<th align='left' rowspan=2>Keterangan</th>
	 	<th colspan='3'>ARM</th>
	 	<th colspan='3'>R30</th>
	 	<th colspan='3'>BRE</th>
	 	<!--<th colspan='3'>BIA</th>
	 	<th colspan='3'>KAA</th>-->
	 </tr>	
	 <tr>
		
		<th>HI</th>
		<th>SHI</th>
		<th>SBI</th>
		<th>HI</th>
		<th>SHI</th>
		<th>SBI</th>
		<th>HI</th>
		<th>SHI</th>
		<th>SBI</th>
		<!--<th>HI</th>
		<th>SHI</th>
		<th>SBI</th>
		<th>HI</th>
		<th>SHI</th>
		<th>SBI</th>-->
		
	 </tr>
	
	 <tr>
		<tbody id="bodytable">
		</tbody>
	</tr>
</table>
<br>
<br>
<!--<table class="table2" border=1 >
	 <tr>
	 	<th align='left'>Keterangan</th>
	 	<th>ARM SHI</th>
	 	<th>R30 SHI</th>
	 	<th>BRE SHI</th>
	 	<th>BIA SHI</th>
	 	<th>KAA SHI</th>
	 </tr>	
	
	 <tr>
		<tbody id="bodytableshi">
		</tbody>
	</tr>

</table>-->
</div>
</body>
</html>




