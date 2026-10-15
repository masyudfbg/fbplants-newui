<!--#include file="txtFunction.inc"-->

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
    width: 80%;
    border: 1px solid #f2f5f7;
	font-size: 10px;
	
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
	width:150px;
}
</style>
    </head>
	
<%
dim company
company    = Request("company")
%>

<script src="/js/jquery-1.9.1.js"></script>	
<script>
	$(function() {
	var txtComp = $.urlParam('company');
		getMonitoring(txtComp);
	});
	
	$.urlParam = function(name){
    var results = new RegExp('[\?&]' + name + '=([^&#]*)').exec(window.location.href);
    if (results==null) {
       return null;
    }
    return decodeURI(results[1]) || 0;
	}

	function getMonitoring(strcomp)
	{			$('#bodytable').empty();
				$.post("GetMonitoringDataInputView.asp?company="+strcomp+"&validasi=" + 0, function(data) {					
					var json  = JSON.parse(data);
					var jumlviewtbs  = json.dataview.length;
					
					for (var i = 0; i < jumlviewtbs; i++) {
						
						var day_4   = json.dataview[i].hmin4_point;
						var day_3 	= json.dataview[i].hmin3_point;
						var day_2 	= json.dataview[i].hmin2_point;
						var day_1 	= json.dataview[i].hmin1_point;
						if (jumlviewtbs>0){
							var docketlist = '';
							
							docketlist += '<tr>';
							docketlist += '<td align="left">' + json.dataview[i].fccompanycode + '</td>';
							docketlist += '<td align="left">' + json.dataview[i].fcba_name + '</td>';
							docketlist += '<td align="left">' + json.dataview[i].id + '</td>';
							docketlist += '<td align="left">' + json.dataview[i].kodelink + '</td>';
							
							
							if (day_4 == 0)  {
							docketlist += '<td bgcolor= "Red" align="right">' + numformat2(day_4) + '%</td>';
							}
							else {
							docketlist += '<td bgcolor= "Green"  align="right">' + numformat2(day_4) + '%</td>';}
							
							if (day_3 == 0)  {
							docketlist += '<td bgcolor= "Red" align="right">' + numformat2(day_3) + '%</td>';
							}
							else {
							docketlist += '<td bgcolor= "Green"  align="right">' + numformat2(day_3) + '%</td>';}
							
							if (day_2 == 0)  {
							docketlist += '<td bgcolor= "Red" align="right">' + numformat2(day_2) + '%</td>';
							}
							else {
							docketlist += '<td bgcolor= "Green"  align="right">' + numformat2(day_2) + '%</td>';}
							
							if (day_1 == 0)  {
							docketlist += '<td bgcolor= "Red" align="right">' + numformat2(day_1) + '%</td>';
							}
							else {
							docketlist += '<td bgcolor= "Green"  align="right">' + numformat2(day_1) + '%</td>';}
							
							docketlist += '</tr>';
							$('#bodytable').append(docketlist);
						}
						
					}
				});

	}

	

	// Format tampilan tabel: ribuan koma + 2 desimal (1,234,567.89). numformat() lama tetap dipakai untuk input.
	function numformat2(val) {
		var n = parseFloat(String(val == null ? '' : val).replace(/,/g, ''));
		if (isNaN(n)) return val == null ? '' : val;
		var p = (Math.abs(n) < 0.005 ? 0 : n).toFixed(2).split('.');
		return p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + p[1];
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
dim tglHariIni
tglHariIni = textdate(Now)

%>
<br>
<div align="center">
	<span style="font-size:22px;"><strong>MONITORING DATA INPUTAN</strong></span>
</div>
<div align="center">
	<span style="font-size:14px;"><strong>Tanggal Hari Ini: <%=tglHariIni%></strong></span>
</div>

<div align="center">
<%
Dim jmHari
jmHari = cDay(Now)+5

%>
<table class="table1" border=1  width="85%" >
	 <tr>
	 	<th align='left' rowspan=2>PT</th>
		<th align='left' rowspan=2>Unit</th>
		<th align='left' rowspan=2>Modul</th>
		<th align='left' rowspan=2>Kode</th>
		<th colspan='<%=jmHari+1%>'>Tanggal</th>
	 </tr>	
	 <tr>
	 <% for i=4 to 1 Step -1 %>
	 	<th><%=cDay(textdate(Now-i))%></th>
	 <% next %>
	 </tr>
	
	 <tr>
		<tbody id="bodytable">
		</tbody>
	</tr>
</table>
<br>
<br>

</div>
</body>
</html>




