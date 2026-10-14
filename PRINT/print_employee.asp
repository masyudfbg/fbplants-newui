<!--#include file="../connDb1.Inc"-->
<!--#include file="../txtFunction.inc"-->
<html lang="en">
<!--<![endif]-->
<head>
	<meta charset="utf-8" />
	<title>Color Admin | Invoice</title>
	<meta content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" name="viewport" />
	<meta content="" name="description" />
	<meta content="" name="author" />
	
	<!-- ================== BEGIN BASE CSS STYLE ================== -->
	<link href="http://fonts.googleapis.com/css?family=Open+Sans:300,400,600,700" rel="stylesheet" />
	<link href="assets/plugins/jquery-ui/jquery-ui.min.css" rel="stylesheet" />
	<link href="assets/plugins/bootstrap/4.1.0/css/bootstrap.min.css" rel="stylesheet" />
	<link href="assets/plugins/font-awesome/5.0/css/fontawesome-all.min.css" rel="stylesheet" />
	<link href="assets/plugins/animate/animate.min.css" rel="stylesheet" />
	<link href="assets/css/default/style.min.css" rel="stylesheet" />
	<link href="assets/css/default/style-responsive.min.css" rel="stylesheet" />
	<link href="assets/css/default/theme/default.css" rel="stylesheet" id="theme" />
	<!-- ================== END BASE CSS STYLE ================== -->
	
	<!-- ================== BEGIN PAGE LEVEL CSS STYLE ================== -->
	<link href="assets/css/default/invoice-print.min.css" rel="stylesheet" />
	<!-- ================== END PAGE LEVEL CSS STYLE ================== -->
	
	<!-- ================== BEGIN BASE JS ================== -->
	<script src="assets/plugins/pace/pace.min.js"></script>
	<!-- ================== END BASE JS ================== -->
</head>
<body style= "font-size:13px">
<%
	queryektp = "select emp.*, 'PT.' || bu.fccompanyname fccompanyname, sysdate printdate "&_
				" from employee emp"&_
				" inner join businessunit bu on  emp.fcba = bu.fccode "&_
				" where emp.fcba = '"&Request("fcba")&"' and emp.fccode = '"&Request("fccode")&"'"
	'response.write "queryektp : " & queryektp & "<br/>"
	set rssektp = Server.CreateObject("ADODB.Recordset")
	rssektp.Open queryektp, connstr
	if not rssektp.eof then
%>		
			<!-- begin breadcrumb -->
			<ol class="breadcrumb hidden-print pull-right">
				<li class="breadcrumb-item"><a href="javascript:;" onclick="konvert()" class="btn btn-sm btn-white m-b-10 p-l-5"><i class="fa fa-print t-plus-1 fa-fw fa-lg"></i> Export as Pdf</a></li>
			</ol>
			<!-- end breadcrumb -->
			<!-- begin page-header -->
			<h1 class="page-header hidden-print">Print Employee <small></small></h1>
			<!-- end page-header -->

			<!-- begin invoice -->
			<div id="invoice" class="invoice">
				<div class="col-lg-12 center ui-sortable" style="float:none;margin:auto;">
                <!-- begin invoice-company -->
                <div style="height:70px;">
                    <img src="../iplas.png" style="float:left;" width="105" />
                </div>
                <!-- end invoice-company -->
                <!-- begin invoice-header -->
                <div class="invoice-header">
                    <div class="invoice-from">
                        <small></small>
                        <address class="m-t-5 m-b-5">
                            <strong class="text-inverse"><%=rssektp("fccompanyname") %></strong><br />
                            <div class="date text-inverse m-t-5"><b><%=rssektp("fcba") %></b></div>
                        </address>
                    </div>
                    <div class="invoice-to">
                       
                    </div>
                    <div class="invoice-date">
                        <small>Tanggal Print</small>
                        <div class="date text-inverse m-t-5"><%=rssektp("printdate") %></div>
                    </div>
                </div>
                <!-- end invoice-header -->
                <!-- begin invoice-content -->
                <div class="invoice-content">
                	<!-- begin table-responsive -->
                    <div class="table-responsive">
                        <table class="table table-invoice">
                            <tbody>								
								<tr>
                                    <td width="19%">Kebutuhan</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left">
										<input type="checkbox">&nbsp;Baru</input>&nbsp;
										<input type="checkbox">&nbsp;Perubahan</input>&nbsp;
										<input type="checkbox">&nbsp;Mutasi</input>&nbsp;
										<input type="checkbox">&nbsp;Hapus</input>&nbsp;
									</td>
                                    <td width="19%" class="text-left"></td>
                                    <td width="1%" class="text-center"></td>
									<td width="30%" class="text-left"></td>
                                </tr>
								<tr>
                                    <td width="19%">Deskripsi</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left" colspan="4">
										&nbsp;
									</td>                                   
                                </tr>
								<tr>
                                    <td width="19%" rowspan="2">&nbsp;</td>
                                    <td width="1%" rowspan="2" class="text-center">&nbsp;</td>
									<td width="30%" class="text-left" colspan="4">
										&nbsp;
									</td>                                   
                                </tr>
								<tr>
                                    
                                    
									<td width="30%" class="text-left" colspan="4">
										&nbsp;
									</td>                                   
                                </tr>
                                <tr>
                                    <th colspan="6" bgcolor="#e2e7eb">Personal Information</th>
                                </tr>							
                                <tr>
                                    <td width="19%">ID Karyawan</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("fccode") %></td>
                                    <td width="19%" class="text-left">Nama Karyawan</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("fcname") %></td>
                                </tr>
								<tr>
                                    <td width="19%">Tanggal Lahir</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("dateofbirth") %></td>
                                    <td width="19%" class="text-left">Agama</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("religion") %></td>
                                </tr>    
								<tr>
                                    <td width="19%">Suku</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("race") %></td>
                                    <td width="19%" class="text-left">Kewarganegaraan</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("nationality") %></td>
                                </tr> 
								<tr>
                                    <td width="19%">Jenis Kelamin</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("gender") %></td>
                                    <td width="19%" class="text-left">Alamat</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("address") %></td>
                                </tr>   
								<tr>
                                    <td width="19%">Pendidikan</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("highesteducation") %></td>
                                    <td width="19%" class="text-left">Tahun Kelulusan</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("yearofgraduation") %></td>
                                </tr>
								<tr>
                                    <td width="19%">Status Perkawinan</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("Maritalstatus") %></td>
                                    <td width="19%" class="text-left">Nama Pasangan</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("spouse") %></td>
                                </tr>
								<tr>
                                    <td width="19%">Jumlah Anak</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("Dependentchild") %></td>
                                    <td width="19%" class="text-left">E-KTP</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("identitycard") %></td>
                                </tr> 
								<tr>
                                    <th colspan="6" bgcolor="#e2e7eb">Job Information</th>
                                </tr>							
                                <tr>
                                    <td width="19%">Divisi</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("sectionname") %></td>
                                    <td width="19%" class="text-left">Kelompok</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("gangcode") %></td>
                                </tr>
								<tr>
                                    <td width="19%">Tgl Bergabung</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("dateofjoin") %></td>
                                    <td width="19%" class="text-left">Tgl Pengangkatan</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("dateofdesignation") %></td>
                                </tr>
								<tr>
                                    <td width="19%">Status Karyawan</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("presentgrade") %></td>
                                    <td width="19%" class="text-left">Gaji Pokok</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("Presentrate") %></td>
                                </tr>
								<tr>
                                    <td width="19%">Posisi</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("position") %></td>
                                    <td width="19%" class="text-left">Tunjangan Beras</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("naturetype") %></td>
                                </tr>
								<tr>
                                    <td width="19%">Golongan</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("subgrade") %></td>
                                    <td width="19%" class="text-left">ID Perusahaan</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("company_nik") %></td>
                                </tr>
								<tr>
                                    <td width="19%">Tanggal Keluar</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("dateterminate") %></td>
                                    <td width="19%" class="text-left"></td>
                                    <td width="1%" class="text-center"></td>
									<td width="30%" class="text-left"></td>
                                </tr>
								<tr>
                                    <th colspan="6" bgcolor="#e2e7eb">BPJS Information</th>
                                </tr>							
                                <tr>
                                    <td width="19%">BPJS TK</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("jamsostekcategory") %></td>
                                    <td width="19%" class="text-left">No BPJS TK</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("jamsostekno") %></td>
                                </tr>
								<tr>
                                    <td width="19%">BPJS Kes</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("bpjskes_category") %></td>
                                    <td width="19%" class="text-left">No BPJS Kes</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("bpjskes_no") %></td>
                                </tr>
								<tr>
                                    <th colspan="6" bgcolor="#e2e7eb">Tax Information</th>
                                </tr>							
                                <tr>
                                    <td width="19%">Status Pajak</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("includedtaxcalculation") %></td>
                                    <td width="19%" class="text-left">No NPWP</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("npwp") %></td>
                                </tr>
								<tr>
                                    <td width="19%">Status Pajak Keluarga</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("taxstatus") %></td>
                                    <td width="19%" class="text-left">Tanggungan perusahaan</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("taxpaidbycompany") %></td>
                                </tr>
								<tr>
                                    <th colspan="6" bgcolor="#e2e7eb">Bank Information</th>
                                </tr>							
                                <tr>
                                    <td width="19%">Tipe Penggajian</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("paymenttype") %></td>
                                    <td width="19%" class="text-left">Nama Bank</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("bankname") %></td>
                                </tr>
								<tr>
                                    <td width="19%">No Rekening</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("accountno") %></td>
                                    <td width="19%" class="text-left">Nama pemilik rekening</td>
                                    <td width="1%" class="text-center">:</td>
									<td width="30%" class="text-left"><%=rssektp("ownershipofaccount") %></td>
                                </tr>
								<tr>
                                    <th colspan="6" >&nbsp;</th>
                                </tr>
								<tr>
                                    <td  style="border-color:white;" width="19%">Persetujuan :</td>
                                    <td  style="border-color:white;" width="1%" class="text-left"></td>
									<td  style="border-color:white;" width="30%" class="text-left">Dibuat Oleh,</td>
                                    <td  style="border-color:white;" width="19%" colspan="2" class="text-left">Diketahui oleh,</td>
									<td  style="border-color:white;" width="30%" class="text-center">Disetujui oleh,</td>
                                </tr>
								<tr>
                                    <th  style="border-color:white;" colspan="6" >&nbsp;</br></br></br></br></th>
                                </tr>
								<tr>
                                    <td style="border-color:white;" width="19%"></td>
                                    <td style="border-color:white;" width="1%" class="text-left"></td>
									<td style="border-color:white;" width="30%" class="text-left">(--------------------)</td>
                                    <td style="border-color:white;" width="19%" colspan="2" class="text-left">(--------------------)</td>
									<td style="border-color:white;" width="30%" class="text-center">(--------------------)</td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                	<!-- end table-responsive -->
                </div>
                <!-- end invoice-content -->
                <!-- begin invoice-note -->
                <div class="invoice-note">
                     * Mohon dokumen ini dapat digunakan sebaik - baiknya<br /> 
                </div>
                <!-- end invoice-note -->
                <!-- begin invoice-footer -->
                <div class="invoice-footer">
                    <p class="text-left m-b-5 f-w-600">
                        &nbsp
                    </p>
                    <p class="text-center">
                        <span class="m-r-10"><i class="fa fa-fw fa-lg fa-globe"></i> amsteknologi.com</span> 	 	 	
                        <span class="m-r-10"><i class="fa fa-fw fa-lg fa-phone-volume"></i> T:021-21889988</span>
                        <span class="m-r-10"><i class="fa fa-fw fa-lg fa-envelope"></i> functional.ho@amsteknologi.com</span>
                    </p>
                </div>
                <!-- end invoice-footer -->
				</div>
            </div>
			<!-- end invoice -->
<%
	end if
	rssektp.close()
	rssektp=null
%>	
	<!--Tambahkan Libraries - JQuery and jspdf-->
	<script src="https://code.jquery.com/jquery-1.12.3.min.js"></script>
	<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/1.4.0/jspdf.min.js"></script>
	<script src="https://cdnjs.cloudflare.com/ajax/libs/rasterizehtml/1.3.0/rasterizeHTML.allinone.js"></script>
	<script type="text/javascript">
		var doc = new jsPDF('p', 'pt', 'a4');
			var specialElementHandlers = {
				'#editor': function (element, renderer) {
				return true;
			}
		};
		
		var w = 1000;
		var h = 1000;
		
		
		$('#konvert').click(function () {   
			//doc.internal.scaleFactor = 5;
			doc.addHTML($('#invoice')[0], function () {
				 doc.save('Test.pdf');
			});
		});
		
		function konvert(){
			html2canvas(document.getElementById("invoice"), {
            scale: 0,
			onrendered: function (canvas) {
			  var imgData = canvas.toDataURL('image/png');					
                var doc = new jsPDF("p", "mm", "a4");
				//var width = doc.internal.pageSize.getWidth();
				//var height = doc.internal.pageSize.getHeight();				
                doc.addImage(imgData, 'PNG', 0, 0, 210, 250);
                doc.save('sample2.pdf');
				}
			});
		}
	</script>
</body>
</html>