<!--#include virtual="/connDb1.inc"-->
<!--#include virtual="/txtFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<title>lembarA1</title>
<link href="/css/font/font-awesome.css" rel="stylesheet">
<script TYPE="text/javascript"> 
<!-- hide 
function killerrors()
{ 
return true; 
} 
window.onerror = killerrors; 
// --> 
</script>
<style type="text/css">
.setting_content{float:left;margin-left:20px;}
.image { 
   position: relative; 
   width: 100%; /* for IE 6 */
}
*{font-style:normal;
font-weight:normal;
font-size:12px;
font-family:Arial;}
.setting{
	position:fixed;
	background-color:rgba(12,91,36,0.9);
	width:100%;
	min-height:20px;
	top:0px;
	left:0px;
	right:0px;
	padding:5px;
	z-index:1000;
	color:#FFFFFF;
	box-shadow:1px 1px 8px rgba(0,0,0,0.4);
	-o-box-shadow:1px 1px 8px rgba(0,0,0,0.4);
	-moz-box-shadow:1px 1px 8px rgba(0,0,0,0.4);
	-webkit-box-shadow:1px 1px 8px rgba(0,0,0,0.4);
}
.watermark_number{
	position:absolute;
	top:30%;
	width:100%;
	text-align:center;
	color:rgba(0,0,0,0.3);
}
</style>
</head>
<script TYPE="text/javascript">
sessionFCBA = "<%=Session("FCBA")%>";
if (sessionFCBA === ""){
	alert("Silahkan Login Kembali");
}

var currentpos,timer; 
function initialize() 
{ 
timer=setInterval("scrollwindow()",10);
} 
function sc(){
clearInterval(timer); 
}
function scrollwindow() 
{ 
currentpos=document.body.scrollTop; 
window.scroll(0,++currentpos); 
if (currentpos != document.body.scrollTop) 
sc();
} 
document.onmousedown=sc
document.ondblclick=initialize

</script>
<script> 
	function format_money_INA(n, currency) {
		return currency + " " + n.toFixed(0).replace(/./g, function(c, i, a) {
			return i > 0 && c !== "." && (a.length - i) % 3 === 0 ? "." + c : c;
		});
	}
	function format_money_ING(n, currency) {
		return currency + " " + n.toFixed(0).replace(/./g, function(c, i, a) {
			return i > 0 && c !== "." && (a.length - i) % 3 === 0 ? "," + c : c;
		});
	}
	
	function harusangka(jumlah){
		var karakter = (jumlah.which) ? jumlah.which : event.keyCode
		if (karakter > 31 && (karakter < 48 || karakter > 57))
		return false;
		
	return true;
}
</script>
<style>

div.image {
	page:image;
}
div.box{
	border:1px #000000 solid;
	width:22px;
	height:16px;
	float:left;
	text-align:center;
	margin-left:-1px;
	margin-bottom:2px;
	font-style:normal;
	font-weight:normal;
	font-size:12px;
	font-family:Arial;
	color:#000000;
	padding-top:2px;

}
div.box_space{
	border:1px #000000 solid;
	width:22px;
	height:16px;
	float:left;
	text-align:center;
	margin-left:-1px;
	margin-bottom:2px;
	margin-right:23px;
	font-style:normal;
	font-weight:normal;
	font-size:12px;
	font-family:Arial;
	color:#000000;
	padding-top:2px;
}


@page {
	size: A4;
    margin: 0.5cm auto;
}
@media screen {
	html{
		background-color:#333333;
	}
	body{
	padding-top:40px;
	background-color:rgba(255,255,255,0);
	}
	#canvas_{
		WIDTH:870px;
	}
	.image {
	margin-top:10px;
	both:clear;
	box-shadow:1px 0px 3px rgba(0,0,0,0.4);
	-o-box-shadow:1px 0px 3px rgba(0,0,0,0.4);
	-moz-box-shadow:1px 0px 3px rgba(0,0,0,0.4);
	-webkit-box-shadow:1px 0px 3px rgba(0,0,0,0.4);
	}
	.watermark_number{
		display:block;
	}
}

@media print {
	html{
		background-color:#FFFFFF;
	}
	body{
	padding-top:0px;
	background-color:rgba(255,255,255,0);
	}
	.image {
	page-break-after: always;
	margin-top:3px;
	both:clear;
	box-shadow:none;
	-o-box-shadow:none;
	-moz-box-shadow:none;
	-webkit-box-shadow:none;
	}
    
	#canvas_{
		WIDTH:870px;
	}
	.setting{
		display:none;
	}
	.watermark_number{
		display:none;
	}
}
#Content a:after { /* select links in the content area only */
      content: " [" attr(href) "] ";
      display: none;
    }
</style>  
<%

dim txtYear,txtFCBA

txtYear = Year(Now) - 1
txtFCBA = Session("FCBA")
txtFILTERHAL = 2
'// ============= FUNGSI ARRAY ============= //
		Function GetRSArray(strSQL)
			set objRSArray = connStr.Execute(strSQL)
			If Not objRSArray.EOF Then
				arrRS = objRSArray.GetRows()
				GetRSArray = arrRS
			Else
                Exit Function
			End If
			objRSArray.Close()
			Set objRSArray = Nothing
		End Function
	
'// ============= Query ============= //'Apabila ada penambahan Kolom, lakukan di bawah #LCT	
stringUSER  = 	"select businessunit_api.get_NPWPOWNERCOMPANY(a.fcba) NPWPOWNERCOMPANY," & _
				"	businessunit_api.get_NPWPCOMPANY(a.fcba) NPWPCOMPANY," & _
				"	businessunit_api.get_NPWPOWNERDIRECTOR(a.fcba) NPWPOWNERDIRECTOR," & _
				"	businessunit_api.get_NPWPDIRECTOR(a.fcba) NPWPDIRECTOR," & _
				"	'Tanjung Simpang' area," & _
				"	   b.fccode," & _
				"	   b.fcname," & _
				"	   b.gender," & _
				"	   b.npwp," & _
				"	   b.taxstatus," & _
				"	   empposition_api.get_description(b.position,a.fcba) position," & _
				"	   b.dateofjoin," & _
				"	   b.address," & _
				"	   a.* " & _
				"  from EMP_YEARLY_INCOME a, employee b" & _
				" where b.fccode = a.empcode " & _
				"   and b.fcba = a.fcba " & _
				"   and a.fcba = '" & txtFCBA & "' " & _
				"   and a.year = '" & txtYear & "' " & _
				"   and b.includedtaxcalculation='TRUE' " & _
				" order by a.empcode " 

'response.write(stringUSER)	
'				"   and b.npwp is not null " & _		
arrRS = GetRSArray(stringUSER)
%>		
<script>
halaman_awal = <%=LBound(arrRS, 2)%>;
jmlHalaman   = <%=UBound(arrRS, 2)%>;
</script>
<script> 

//isi Form A1
date_now 					= new Date();
TAHUN_cetak					= <%=txtYear%>;

// ========================== DATA PERUSAHAAN ========================== //
nama_perusahaan_array		= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) %>
								'<%=arrRS(0, i)%>',
								<%
								Next
								End If 
								%>];
NPWP_perusahaan_array		= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2)
									if arrRS(1, i) <> "" then 
										npwp = replace(arrRS(1, i),".","")
										%>
										'<%=replace(npwp,"-","")%>',
										<%
									else %>
									'',
									<% 
									end if
								Next	
								End If 
								%>];
NPWP_ownerdirector_array	= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) %>
								'<%=arrRS(2, i)%>',
								<%
								Next	
								End If 
								%>];
NPWP_director_array			= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2)
									if arrRS(3, i) <> "" then 
										npwp = replace(arrRS(3, i),".","")
										%>
										'<%=replace(npwp,"-","")%>',
										<%
									else %>
									'',
									<% 
									end if
								Next	
								End If 
								%>];								
tempat_perusahaan_array		= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) %>
								'<%=arrRS(4, i)%>',
								<%
								Next	
								End If 
								%>];
// ========================== DATA OWNER ========================== //
no_urut_isi_array 			= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) %>
								'<%=arrRS(5, i)%>',
								<%
								Next
									
								End If 
								%>];
nama_pegawai_array 			= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) %>
								'<%=arrRS(6, i)%>',
								<%
								Next
									
								End If 
								%>];
gender_array				= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) %>
								'<%=arrRS(7, i)%>',
								<%
								Next	
								End If 
								%>];
NPWP_pegawai_array			= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
									if arrRS(8, i) <> "" then 
										npwp = replace(arrRS(8, i),".","")
										%>
										'<%=replace(npwp,"-","")%>',
										<%
									else %>
									'',
									<% 
									end if
								  Next
								End If 
								%>];
alamat_pegawai_array		= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								'<%=arrRS(12, i)%>',
								<%
								Next
								End If 
								%>];
maritalstatus_array			= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								'<%=arrRS(9, i)%>',
								<%
								Next
								End If 
								%>];
jabatan_array				= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								'<%=arrRS(10, i)%>',
								<%
								Next
								End If 
								%>];
masa_perolehan_array		= [<% If IsArray(arrRS) Then
								dim  month_sp
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								
										if Year(arrRS(11, i)) < txtYear then
											month_sp = "01"	
										elseif Year(arrRS(11, i)) = txtYear then
											if len(Month(arrRS(11, i))) = 1 then
												month_sp = "0" & Month(arrRS(11, i))	
											else
												month_sp = Month(arrRS(11, i))	
											end if
										end if
										
								%>
										'<%=month_sp%>',
								<%
									Next
								End If 
								%>];
masa_perolehan_SD_array	= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
										'12',
								<%
									Next
								End If 
								%>];

// ========================== PENGHASILAN BRUTO ========================== //
gaji_1_array				= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								<%=arrRS(16, i)%>,
								<%
								Next
								End If 
								%>]; 	// GAJI/PENSIUN ATAU THT / JHT
tunjangan_pph_2_array		= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								0,
								<%
								Next
								End If 
								%>];	// TUNJANGAN PPh
tunjangan_lainya_3_array	= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								0,
								<%
								Next
								End If 
								%>];	// TUNJANGAN LAINYA, UANG LEMBUR, DAN SEBAGAINYA
honorarium_4_array			= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								0,
								<%
								Next
								End If 
								%>];
premiasuransi_5_array		= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								0,
								<%
								Next
								End If 
								%>];
penerimaan_natura_6_array	= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								0,
								<%
								Next
								End If 
								%>];
tbgjt_array					= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								0,
								<%
								Next
								End If 
								%>];

// ========================== PENGURANGAN ========================== //
biaya_jabatan_7_array		= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								<%=replace(arrRS(17, i),"-","")%>,
								<%
								Next
								End If 
								%>]; // BIAYA JABATAN / BIAYA PENSIUN ATAS PPENGHASILAN PADA ANGKA 7
biaya_jabatan_8_array		= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								0,
								<%
								Next
								End If 
								%>]; // BIAYA JABATAN / BIAYA PENSIUN ATAS PPENGHASILAN PADA ANGKA 8
iuran_pensiun_array			= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								0,
								<%
								Next
								End If 
								%>]; // IURAN PENSUN ATAU IURAN THT/JHT

// ========================== PERHITUNGAN PPh Pasal 21 ========================== //
penghasilan_sebelum_array	= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								0,
								<%
								Next
								End If 
								%>];
jml_penghasilan_neto_array	= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								<%=arrRS(18, i)%>,
								<%
								Next
								End If 
								%>];
penghasilan_tdk_pjk_array	= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								<%=arrRS(19, i)%>,
								<%
								Next
								End If 
								%>];
pph_kn_pjk_array			= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								<%=replace(arrRS(22, i),"-","")%>,
								<%
								Next
								End If 
								%>]; // PPh PASAL 21 ATAS PENGHASILAN KENA PAJAK SETAHUN/DISETHUNKAN
pph_telah_dipotong_array	= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								0,
								<%
								Next
								End If 
								%>]; // PPh PASAL 21 YANG TELAH DI POTONG MASA SEBELUMNYA
pph_terhutang_array			= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								<%=replace(arrRS(22, i),"-","")%>,
								<%
								Next
								End If 
								%>]; // PPh PASAL 21 TERHUTANG
pph2126_tlh_diptng_array	= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								<%=replace(arrRS(22, i),"-","")%>,
								<%
								Next
								End If 
								%>]; // PPh PASAL 21 DAN PPh PASAL 26 YANG TELAH DIOTONG DAN DILUNASI
ssp_pph_pasal21_array		= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								<%=replace(arrRS(22, i),"-","")%>,
								<%
								Next
								End If 
								%>]; // Dipotong dan di lunasi dengan SSP PPh pasal 21 Ditanggung
potong_dgn_ssp_array		= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								<%=replace(arrRS(22, i),"-","")%>,
								<%
								Next
								End If 
								%>]; // Dipotong dan dilunasi degan SSP


month_21_24_array 			= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								
										if Year(arrRS(11, i)) < txtYear then
											month_sp = "01"	
										elseif Year(arrRS(11, i)) = txtYear then
											if len(Month(arrRS(11, i))) = 1 then
												month_sp = "0" & Month(arrRS(11, i))	
											else
												month_sp = Month(arrRS(11, i))	
											end if
										end if
										
								%>
										'<%=month_sp%>',
								<%
									Next
								End If 
								%>];
year_21_24_array			= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
									<%=year(arrRS(11, i))%>,
								<%
									Next
								End If 
								%>];
jml_pph_pasal_21_24_array 	= [<% If IsArray(arrRS) Then
								For i = LBound(arrRS, 2) To UBound(arrRS, 2) 
								%>
								<%=replace(arrRS(22, i),"-","")%>,
								<%
								Next
								End If 
								%>];

//============================== Jumlah PPh pasal 21 ========================//
jml_pph_pasal_21_array	= "b";
//============================== Jumlah tersebut pada angka 23 telah  ========================//
jml_pph_pasal_24_array	= "a";

</script>
<body id="canvas_" style="margin:auto;" vlink="#FFFFFF" link="#FFFFFF" bgcolor="#ffffff">
	<div class="setting">
		<div class="setting_content" >NAME&nbsp;:&nbsp;&nbsp;
		<select id="halaman" onchange="halaman()">
		<option value="" selected>ALL</option>
		<% If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2) %>
			<option value='number<%=i%>'><%=arrRS(6, i)%></option>
		<%
			Next
		End If 
		%>
		</select>
		</div>
		<div class="setting_content">
			PAGE&nbsp;:&nbsp;&nbsp;
			<input id="from_page" name="from_page" type="text" value="<%=(LBound(arrRS, 2)+1)%>" style="width:50px;text-align:center;" onkeyup="halaman2()" onkeypress="return harusangka(event);" onchange="clear_value('from_page','<%=(LBound(arrRS, 2)+1)%>')">
			&nbsp;&nbsp;TO&nbsp;&nbsp;
			<input id="to_page" name="to_page" type="text" value="<%=(UBound(arrRS, 2))%>" style="width:50px;text-align:center;" onkeyup="halaman2()" onkeypress="return harusangka(event);" onchange="clear_value('to_page','<%=(UBound(arrRS, 2))%>')">
		</div>
		<a id="print" style="cursor:pointer;margin-right:25px;position:absolute;right:1.2%;z-index:100;padding:3px;" onClick="window.print() ">
		<i class="fa fa-print"></i>
		&nbsp;&nbsp;Print
		</a>
	</div>

<script>
for (l = halaman_awal; l <= jmlHalaman; l++) { 

	if(l < 100 ){
		fz_watermark = 500;
	}else if(l >= 100 && l < 1000){
		fz_watermark = 400;
	}else if(l >= 1000 && l < 10000){
		fz_watermark = 300;
	}else{
		fz_watermark = 100;
	}

	document.write('<div id="number' + l + '" class="image">');
	document.write('<div class="watermark_number" style="font-size:' + fz_watermark + 'px;">' + (l + 1) + '</div>');
	
	no_urut_isi 		= no_urut_isi_array[l];
	nama_pegawai 		= nama_pegawai_array[l].toUpperCase();
	gender				= gender_array[l];
	NPWP_pegawai		= NPWP_pegawai_array[l];
	alamat_pegawai		= alamat_pegawai_array[l].toUpperCase();
	maritalstatus		= maritalstatus_array[l];

	jabatan				= jabatan_array[l].toUpperCase();
	masa_perolehan		= masa_perolehan_array[l];
	masa_perolehan_SD	= masa_perolehan_SD_array[l];
	//---------- DATA PERUSAHAAN -----------------------------------------------//
	nama_perusahaan		= nama_perusahaan_array[l].toUpperCase();
	NPWP_perusahaan		= NPWP_perusahaan_array[l];
	tempat_perusahaan	= tempat_perusahaan_array[l].toUpperCase();
	NPWP_ownerdirector	= NPWP_ownerdirector_array[l].toUpperCase();
	NPWP_director		= NPWP_director_array[l];
	// --------- PENGHASILAN BRUTO -------------------------------------------- //
	gaji_1				= gaji_1_array[l];	// GAJI/PENSIUN ATAU THT / JHT
	tunjangan_pph_2		= tunjangan_pph_2_array[l];	// TUNJANGAN PPh
	tunjangan_lainya_3	= tunjangan_lainya_3_array[l];	// TUNJANGAN LAINYA, UANG LEMBUR, DAN SEBAGAINYA
	honorarium_4		= honorarium_4_array[l];
	premiasuransi_5		= premiasuransi_5_array[l];
	penerimaan_natura_6	= penerimaan_natura_6_array[l];
	jml_1_6				= gaji_1 + tunjangan_pph_2 + tunjangan_lainya_3 + honorarium_4 + premiasuransi_5 + penerimaan_natura_6;
	tbgjt				= tbgjt_array[l];
	jml_1_8				= jml_1_6 + tbgjt;
	// --------- PENGURANGAN -------------------------------------------- //
	biaya_jabatan_7		= biaya_jabatan_7_array[l]; // BIAYA JABATAN / BIAYA PENSIUN ATAS PPENGHASILAN PADA ANGKA 7
	biaya_jabatan_8		= biaya_jabatan_8_array[l]; // BIAYA JABATAN / BIAYA PENSIUN ATAS PPENGHASILAN PADA ANGKA 8
	iuran_pensiun		= iuran_pensiun_array[l]; // IURAN PENSUN ATAU IURAN THT/JHT
	jml_pengurangan		= biaya_jabatan_7 + biaya_jabatan_8 + iuran_pensiun;
	// --------- PERHITUNGAN PPh Pasal 21 ------------------------------- //
	jml_penghasilan		= jml_1_8 - jml_pengurangan;
	penghasilan_sebelum	= penghasilan_sebelum_array[l];
	jml_penghasilan_neto= jml_penghasilan_neto_array[l];
	penghasilan_tdk_pjk	= penghasilan_tdk_pjk_array[l];
	penghasilan_kn_pjk	= jml_penghasilan_neto - penghasilan_tdk_pjk;
	pph_kn_pjk			= pph_kn_pjk_array[l]; // PPh PASAL 21 ATAS PENGHASILAN KENA PAJAK SETAHUN/DISETHUNKAN
	pph_telah_dipotong	= pph_telah_dipotong_array[l]; // PPh PASAL 21 YANG TELAH DI POTONG MASA SEBELUMNYA
	pph_terhutang		= pph_terhutang_array[l]; // PPh PASAL 21 TERHUTANG
	pph2126_tlh_diptng	= pph2126_tlh_diptng_array[l]; // PPh PASAL 21 DAN PPh PASAL 26 YANG TELAH DIOTONG DAN DILUNASI
	ssp_pph_pasal21		= ssp_pph_pasal21_array[l]; // Dipotong dan di lunasi dengan SSP PPh pasal 21 Ditanggung
	potong_dgn_ssp		= potong_dgn_ssp_array[l]; // Dipotong dan dilunasi degan SSP

	status				= maritalstatus.substring(0, (maritalstatus.length - 1 )); 
	jumlahtanggungan	= maritalstatus.substring(status.length, maritalstatus.length);
	
	month_21_24			= month_21_24_array[l];
	year_21_24			= year_21_24_array[l];
	jml_pph_pasal_21_24 = jml_pph_pasal_21_24_array[l];

	if (jml_pph_pasal_21_array === "a" ){
		jml_pph_pasal_21_23 = pph_terhutang - pph2126_tlh_diptng;
	}else if (jml_pph_pasal_21_array === "b"){
		jml_pph_pasal_21_23 = pph2126_tlh_diptng - pph_terhutang;
	}
	
// ========================== TAHUN ==========================
	
	document.write('<div style="font-size:183%;letter-spacing:32px; position:absolute;top:4.2%;right:1.2%;">' + TAHUN_cetak + '</div>');
	
// ==========================  No. URUT ========================== //
	document.write('<div id="no_urut" style="position:absolute;top:7.4%;left:33.7%;">');
		no_urut_text = "";
			for (i = 0; i < 7; i++) {
				if(no_urut_isi.length < 7 ){
					if(i < (7 - no_urut_isi.length)){
						no_urut_text += "<div class='box'></div>";
					}else{
						no_urut_text += "<div class='box'>" + no_urut_isi[i-(7 - no_urut_isi.length)] + "</div>";
					}
				}else{
					no_urut_text += "<div class='box'>" + no_urut_isi[i] + "</div>";
				}
			}
			document.write(no_urut_text);
	
	document.write('</div>');
	
// ==========================  NPWP Pemotong pajak ========================== //
	document.write('<div id="NPWP_pemotong" style="position:absolute;top:9.2%;left:33.7%;">');
		NPWP_text = "";
		if (typeof NPWP_perusahaan != "undefined" && NPWP_perusahaan != null && NPWP_perusahaan.length > 0){
			for (i = 0; i < NPWP_perusahaan.length; i++) {
				if (i===1 || i===4 || i=== 7 || i===8 || i===11){
					NPWP_text += "<div class='box_space'>" + NPWP_perusahaan[i] + "</div>";
				}else{
					NPWP_text += "<div class='box'>" + NPWP_perusahaan[i] + "</div>";
				}
			}
		}else{
			for (i = 0; i < 15; i++) {
				if (i===1 || i===4 || i=== 7 || i===8 || i===11){
					NPWP_text += "<div class='box_space'></div>";
				}else{
					NPWP_text += "<div class='box'></div>";
				}
			}
		}
			document.write(NPWP_text);
	document.write('</div>');
	
// ==========================  Nama Pemotong pajak ========================== //
	document.write('<div id="nama_pemotong" style="position:absolute;top:10.90%;left:33.7%;">');
		nama_pemotong_text = "";
			for (i = 0; i < 23; i++) {
				if (nama_perusahaan.length < 23){
					if(i >= nama_perusahaan.length){
						nama_pemotong_text += "<div class='box'></div>";
					}else{
						nama_pemotong_text += "<div class='box'>" + nama_perusahaan[i] + "</div>";
					}
				}else{
					nama_pemotong_text += "<div class='box'>" + nama_perusahaan[i] + "</div>";
				}
			}
			document.write(nama_pemotong_text);
	document.write('</div>');
	
// ==========================  NAMA PEGAWAI ATAU PENERIMA PENSIUN/THT/JHT ========================== //
	document.write('<div  id="nama_pegawai" style="position:absolute;top:12.60%;left:33.7%;">');
		nama_pegawai_text = "";
			for (i = 0; i < 23; i++) {
				if (nama_pegawai.length < 23){
					if(i >= nama_pegawai.length){
						nama_pegawai_text += "<div class='box'></div>";
					}else{
						nama_pegawai_text += "<div class='box'>" + nama_pegawai[i] + "</div>";
					}
				}else{
					nama_pegawai_text += "<div class='box'>" + nama_pegawai[i] + "</div>";
				}
			}
		document.write(nama_pegawai_text);
	document.write('</div>');
	
// ==========================  NPWP PEGAWAI ATAU PENERIMA PENSIUN/THT/JHT   ========================== //
	document.write('<div id="NPWP_pegawai" style="position:absolute;top:14.28%;left:33.7%;">');
		NPWP_pegawai_text = "";
		if (typeof NPWP_pegawai != "undefined" && NPWP_pegawai != null && NPWP_pegawai.length > 0){
			for (i = 0; i < NPWP_pegawai.length; i++) {
				if (i===1 || i===4 || i=== 7 || i===8 || i===11){
					NPWP_pegawai_text += "<div class='box_space'>" + NPWP_pegawai[i] + "</div>";
				}else{
					NPWP_pegawai_text += "<div class='box'>" + NPWP_pegawai[i] + "</div>";
				}
			}
		}else{
			for (i = 0; i < 15; i++) {
				if (i===1 || i===4 || i=== 7 || i===8 || i===11){
					NPWP_pegawai_text += "<div class='box_space'></div>";
				}else{
					NPWP_pegawai_text += "<div class='box'></div>";
				}
			}
		}
			
		document.write(NPWP_pegawai_text);
	document.write('</div>');
	
// ==========================  ALAMAT PEGAWAI ATAU PENERIMA PENSIUN/THT/JHT   ========================== //
	document.write('<div id="alamat_pegawai" style="position:absolute;top:15.90%;left:33.7%; ">');
		NPWP_pegawai_text = "";
			for (i = 0; i < NPWP_pegawai.length; i++) {
				if (i===1 || i===4 || i=== 7 || i===8 || i===11){
					NPWP_pegawai_text += "<div class='box_space'>" + NPWP_pegawai[i] + "</div>";
				}else{
					NPWP_pegawai_text += "<div class='box'>" + NPWP_pegawai[i] + "</div>";
				}
			}
var ST = alamat_pegawai.substring(0, 23); 
			//var StrippedString = alamat_pegawai.replace(/\s/g,"");
			var ST_space = [];
			var ST_space2 = [];
			
			if (alamat_pegawai.length > 23){
				for (x = 0; x < alamat_pegawai.length; x++) {
					if (alamat_pegawai[x] === " " || alamat_pegawai[x] === /\s/g ){
						if (x < 23 ){
							ST_space.push(x);
						}
						ST_space2.push(x);
					}
				}
			}

				
			// tentukan spasi sebelum enter
			ST_awal 	= ST_space.length - 1;
			ST_akhir 	= ST_space.length;	
			ST_sisa		= ST_space[ST_awal]+1;
			s 			= 23 - ST_space[ST_awal];
			//alert(alamat_pegawai.length);
			alamat_pegawai_text = "";
			for (i = 0; i <= 45; i++) {
				
				if (alamat_pegawai.length < 23){
					if(i >= alamat_pegawai.length){
						alamat_pegawai_text += "<div class='box'></div>";
					}else{
						alamat_pegawai_text += "<div class='box'>" + alamat_pegawai[i] + "</div>";
					}
				}else{
					if ( ST_space2[ST_akhir] > 23 ){
						if ( i >= ST_space[ST_awal] && i < 23  ){
							alamat_pegawai_text += "<div class='box'> </div>";					
						}else if ( i <= ST_space[ST_awal]  ){
							alamat_pegawai_text += "<div class='box'>" + alamat_pegawai[i] + "</div>";
						}
						
						if ( i > 23 && i <= (alamat_pegawai.length + (s-1)) ){
							alamat_pegawai_text += "<div class='box'>" + alamat_pegawai[i-s] + "</div>";
						}else if ( i > alamat_pegawai.length  ){
							alamat_pegawai_text += "<div class='box'></div>";
						}
						
					}else{
						if ( i < 23 ){
							alamat_pegawai_text += "<div class='box'>" + alamat_pegawai[i] + "</div>";
						}
						if ( i >= 23 && ST_space2[ST_akhir] === 23 && i <= (alamat_pegawai.length - 2) ){ 
							alamat_pegawai_text += "<div class='box'>" + alamat_pegawai[i + 1] + "</div>";
						}else if ( i > alamat_pegawai.length - 2 ){
							alamat_pegawai_text += "<div class='box'></div>";
						}
						
					}
				}
				if (i === 22){
						alamat_pegawai_text += "<br/>";
				}
			}
			document.write(alamat_pegawai_text);
			document.write('</div>');
// ========================== STATUS ========================== //
	if (status === "K" ){
			document.write('<div class="status" style="left:34.5%;position:absolute;top:19.3%;font-weight:bold;">X</div>');
			document.write('<div class="jml_tanggungan" style="left:37.2%;position:absolute;top:21.05%;font-weight:bold;">' + jumlahtanggungan + '</div>');
		}else if ( status === "TK"){
			document.write('<div class="status" style="left:45.2%;position:absolute;top:19.3%;font-weight:bold;">X</div>');
			document.write('<div class="jml_tanggungan" style="left:45.2%;position:absolute;top:21.05%;font-weight:bold;">' + jumlahtanggungan + '</div>');
		}

		if (gender === "MALE" ){
			document.write('<div class="gender" style="position:absolute;top:19.3%;font-weight:bold;left:58.5%;">X</div>');
		}else if ( gender === "FEMALE"){
			document.write('<div class="gender" style="position:absolute;top:19.3%;font-weight:bold;left:69%;">X</div>');
		}
		document.write('<div id="jabatan" style="position:absolute;top:22.8%;left:34.5%;font-size:10px;">' + jabatan + '</div>');
		document.write('<div style="position:absolute;top:22.7%;left:71.7%;letter-spacing:15px;">' + masa_perolehan + '</div>');
		document.write('<div style="position:absolute;top:22.7%;left:85%;letter-spacing:15px;">' + masa_perolehan_SD + '</div>');
		
		//-- Penghasilan Bruto --//
		document.write('<div style="position:absolute;top:27.2%;right:3%;">' + format_money_INA(gaji_1,"") + '</div>');
		document.write('<div style="position:absolute;top:28.85%;right:3%;">' + format_money_INA(tunjangan_pph_2,"") + '</div>');
		document.write('<div style="position:absolute;top:30.6%;right:3%;">' + format_money_INA(tunjangan_lainya_3,"") + '</div>');
		document.write('<div style="position:absolute;top:32.45%;right:3%;">' + format_money_INA(honorarium_4,"") + '</div>');
		document.write('<div style="position:absolute;top:34.25%;right:3%;">' + format_money_INA(premiasuransi_5,"") + '</div>');
		
		document.write('<div style="position:absolute;top:36.2%;right:3%;">' + format_money_INA(penerimaan_natura_6,"") + '</div>');
		document.write('<div style="position:absolute;top:38.1%;right:3%;">' + format_money_INA(jml_1_6,"") + '</div>');
		document.write('<div style="position:absolute;top:39.93%;right:3%;">' + format_money_INA(tbgjt,"") + '</div>');
		document.write('<div style="position:absolute;top:41.9%;right:3%;">' + format_money_INA(jml_1_8,"") + '</div>');
		document.write('<div style="position:absolute;top:44.9%;right:3%;">' + format_money_INA(biaya_jabatan_7,"") + '</div>');
		
		document.write('<div style="position:absolute;top:46.9%;right:3%;">' + format_money_INA(biaya_jabatan_8,"") + '</div>');
		document.write('<div style="position:absolute;top:48.9%;right:3%;">' + format_money_INA(iuran_pensiun,"") + '</div>');
		document.write('<div style="position:absolute;top:50.8%;right:3%;">' + format_money_INA(jml_pengurangan,"") + '</div>');
		document.write('<div style="position:absolute;top:54%;right:3%;">' + format_money_INA(jml_penghasilan,"") + '</div>');
		document.write('<div style="position:absolute;top:56%;right:3%;">' + format_money_INA(penghasilan_sebelum,"") + '</div>');
		
		document.write('<div style="position:absolute;top:58.1%;right:3%;">' + format_money_INA(jml_penghasilan_neto,"") + '</div>');
		document.write('<div style="position:absolute;top:60.1%;right:3%;">' + format_money_INA(penghasilan_tdk_pjk,"") + '</div>');
		document.write('<div style="position:absolute;top:62.2%;right:3%;">' + format_money_INA(penghasilan_kn_pjk,"") + '</div>');
		document.write('<div style="position:absolute;top:64.2%;right:3%;">' + format_money_INA(pph_kn_pjk,"") + '</div>');
		document.write('<div style="position:absolute;top:66.2%;right:3%;">' + format_money_INA(pph_telah_dipotong,"") + '</div>');
		
		document.write('<div style="position:absolute;top:68.2%;right:3%;">' + format_money_INA(pph_terhutang,"") + '</div>');
		document.write('<div style="position:absolute;top:70.3%;right:3%;">' + format_money_INA(pph2126_tlh_diptng,"") + '</div>');
		document.write('<div style="position:absolute;top:73%;right:27.5%;">' + format_money_INA(ssp_pph_pasal21,"") + '</div>');
		document.write('<div style="position:absolute;top:75.2%;right:27.5%;">' + format_money_INA(potong_dgn_ssp,"") + '</div>');
		
		// ========================== 23- Jumlah PPh pasal 21 ========================= //
		if (jml_pph_pasal_21_array === "a" ){
			document.write('<div class="sign_23" style="position:absolute;left:11.4%;font-weight:bold;top:77.7%;">X</div>');
		}else if ( jml_pph_pasal_21_array === "b"){
			document.write('<div class="sign_23" style="position:absolute;left:11.4%;font-weight:bold;top:79.2%;">X</div>');
		}
		
		document.write('<div style="position:absolute;top:78.5%;right:3%;">' + format_money_INA(jml_pph_pasal_21_23,"") + '</div>');
		// ========================== END:23- Jumlah PPh pasal 21 ========================== //
		
		// ========================== 24- Jumlah tersebut pada angka 23 telah ========================== //
		sign_24_txt = "";
		month_21_24_txt = "";
		year_21_24_txt = "";
		
		sign_24_txt	+= '<div class="sign_24" style="position:absolute;left:11.4%;font-weight:bold;top:';
		month_21_24_txt	+= '<div class="month_21_24" style="position:absolute;right:46.8%;text-align:right;letter-spacing:15.5px;top:';
		year_21_24_txt	+= '<div class="year_21_24" style="position:absolute;right:31%;text-align:right;letter-spacing:15.5px;top:';
		if (jml_pph_pasal_24_array === "a" ){
			sign_24_txt  	+= '81.75%';
			month_21_24_txt	+= '81.7%';
			year_21_24_txt  += '81.7%';
		}else if ( jml_pph_pasal_24_array === "b"){
			sign_24_txt	 	+= '83.5%';
			month_21_24_txt	+= '83.5%';
			year_21_24_txt	+= '83.5%';
		}
		sign_24_txt		+= ';">X</div>';
		month_21_24_txt	+= ';">' + month_21_24 + '</div>';
		year_21_24_txt	+= ';">' + year_21_24 + '</div>';
		jml_pph_pasal_21_24_txt = '<div style="position:absolute;top:82.6%;right:3%;">' + format_money_INA(jml_pph_pasal_21_24,"") + '</div>';
		document.write(sign_24_txt + month_21_24_txt + year_21_24_txt + jml_pph_pasal_21_24_txt);

		// =========================== TANDA TANGAN DAN CAP PERUSAHAAN ================================== //
		
		if ( date_now.getDate() < 10){
				day_ttd 	= "0" + date_now.getDate();
		}else{
				day_ttd	= date_now.getDate();
			}
		if ( date_now.getMonth() < 10){
			month_ttd 		= "0" + date_now.getMonth();
		}else{
			month_ttd 		= date_now.getMonth();
		}
		pemotong_pajak 		= '<div class="pemotong_pajak" style="position:absolute;bottom:10%;left:6%;font-weight:bold;">X</div>'
		tempat_perusahaan	= '<div class="tempat_perusahaan" style="position:absolute;bottom:11.6%;left:41.8%;font-weight:bold;width:18.5%;text-align:center;">' + tempat_perusahaan + '</div>';
		date_now_ttd		= '<div class="date_now" style="position:absolute;bottom:11.57%;right:31.3%;font-weight:bold;letter-spacing:13px;">' + day_ttd + '</div>';
		month_now_ttd		= '<div class="month_now" style="position:absolute;bottom:11.57%;right:23.3%;font-weight:bold;letter-spacing:13px;">' + month_ttd + '</div>';
		year_now_ttd		= '<div class="year_now" style="position:absolute;bottom:11.57%;right:9.6%;font-weight:bold;letter-spacing:15.5px;">' + date_now.getFullYear() + '</div>';
		
		document.write(pemotong_pajak + tempat_perusahaan + date_now_ttd + month_now_ttd + year_now_ttd );
		
		NPWP_ownerdirector_text = "";
			for (i = 0; i < 20; i++) {
				if (NPWP_ownerdirector.length < 20){
					if(i >= NPWP_ownerdirector.length){
						NPWP_ownerdirector_text += "<div class='box'></div>";
					}else{
						NPWP_ownerdirector_text += "<div class='box'>" + NPWP_ownerdirector[i] + "</div>";
					}
				}else{
					NPWP_ownerdirector_text += "<div class='box'>" + NPWP_ownerdirector[i] + "</div>";
				}
			}
		document.write('<div class="NPWP_ownerdirector" style="position:absolute;bottom:7.4%;left:15.2%;">' + NPWP_ownerdirector_text + '</div>');
		
		NPWP_director_text = "";
			for (i = 0; i < NPWP_director.length; i++) {
				if (i===1 || i===4 || i=== 7 || i===8 || i===11){
					NPWP_director_text += "<div class='box_space'>" + NPWP_director[i] + "</div>";
				}else{
					NPWP_director_text += "<div class='box'>" + NPWP_director[i] + "</div>";
				}
			}
		document.write('<div class="NPWP_director" style="position:absolute;bottom:4.9%;left:15.2%;">' + NPWP_director_text + '</div>');
		
		// ================== END:TANDA TANGAN DAN CAP PERUSAHAAN ===========================//
		document.write('<img src="/IMAGES/pg_0001.jpg" style="WIDTH:100%;" ALT=""></div>');

}
</script>
<script TYPE="text/javascript">
function halaman(){
	var from_page = document.getElementById("from_page");
	var to_page = document.getElementById("to_page");
	var srt_page = document.getElementById("halaman").value;
	
	if (srt_page === "" ){
		for (jml_id = halaman_awal; jml_id <= jmlHalaman; jml_id++){
			var number_id = "number"+jml_id+"";
			var id_names = document.getElementById(number_id);
			if (srt_page === ""){
				id_names.style.display = "block";
			}
		}
		from_page.value 	= 1; 
		to_page.value		= jmlHalaman+1;
	}else{
		
		for (jml_id = halaman_awal; jml_id <= jmlHalaman; jml_id++){
			var number_id = "number"+jml_id+"";
			var id_names = document.getElementById(number_id);
			
			//alert(number_id);
			if (number_id === srt_page){
				id_names.style.display = "block";
				from_page.value 	= jml_id+1; 
				to_page.value		= jml_id+1;
			}else{
				id_names.style.display = "none";
			}
		}
		
	}

	
}
function halaman2(){
	var from_page = (document.getElementById("from_page").value - 1);
	var to_page = (document.getElementById("to_page").value - 1);
	
	if (from_page === "" && to_page === "" ){
		for (i = halaman_awal; i <= jmlHalaman; i++){
			var number_id = "number"+i+"";
			var id_names = document.getElementById(number_id);
			if (srt_page === ""){
				id_names.style.display = "block";
			}
		}
	}else{
		for (i = halaman_awal; i <= jmlHalaman; i++){
			var number_id = "number"+i+"";
			var id_names = document.getElementById(number_id);
			
			//alert(number_id);
			if ( i >= from_page && i <= to_page){
				id_names.style.display = "block";
			}else{
				id_names.style.display = "none";
			}
		}
	}
	
}
function clear_value(val,nrml){
	var id = document.getElementById(val);
	if (id.value >= ( halaman_awal + 1 ) || id.value <= ( jmlHalaman + 1 )){
		id.value = id.value;
	}else{
		id.value = nrml;
	}
}


var srt_page = document.getElementById("halaman").value;
	if (srt_page === "" ){
		for (jml_id = halaman_awal; jml_id <= jmlHalaman; jml_id++){
			var number_id = "number"+jml_id+"";
			var id_names = document.getElementById(number_id);
			if (srt_page === ""){
				id_names.style.display = "block";
			}
		}
	}else{
		
		for (jml_id = halaman_awal; jml_id <= jmlHalaman; jml_id++){
			var number_id = "number"+jml_id+"";
			var id_names = document.getElementById(number_id);
			
			//alert(number_id);
			if (number_id === srt_page){
				id_names.style.display = "block";
			}else{
				id_names.style.display = "none";
			}
		}
	}
	
	
var from_page = (document.getElementById("from_page").value - 1);
	var to_page = (document.getElementById("to_page").value - 1);
	
	if (from_page === "" && to_page === "" ){
		for (i = halaman_awal; i <= jmlHalaman; i++){
			var number_id = "number"+i+"";
			var id_names = document.getElementById(number_id);
			if (srt_page === ""){
				id_names.style.display = "block";
			}
		}
	}else{
		
		for (i = halaman_awal; i <= jmlHalaman; i++){
			var number_id = "number"+i+"";
			var id_names = document.getElementById(number_id);
			
			//alert(number_id);
			if ( i >= from_page && i <= to_page){
				id_names.style.display = "block";
			}else{
				id_names.style.display = "none";
			}
		}
	}
	var currentZoom = parent.ltop.currentZoom;
	if(currentZoom != undefined)
	document.body.style.zoom=currentZoom/100;
	
	
</script>
</body>
</html>