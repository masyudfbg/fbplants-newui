<!--#include file="connDb1.Inc"-->
<!--#include file="txtFunction.inc"-->
<!--#include virtual="/adovbs.inc"-->
<!--#include virtual="/REPORT/functions.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="/css/stylereport2.css" type="text/css" />
<title>ePlaS Reporting - Laporan Harian</title>

<script language="JavaScript" src='/js/x.js'></script>
	<script language="JavaScript" src='/js/xtableheaderfixed.js'></script>
	<script language="JavaScript">
	xAddEventListener(window, 'load',
	  function() {
		new xTableHeaderFixed('xthf-site', window);
	  }, false
	);

	</script>


<style type="text/css">
@media print
{
.printhide{display:none;}
}

.tebal{
	font-weight:bold;
}

.ratakiri{
	text-align:left;
}
.ratakanan{
	text-align:right;
}
.ratatengah{
	text-align:center;
}
.lebarket{
	width:780px;
}

</style>
</head>

<%



Dim tbl,doc,nday,txtSQL,rsdata,tglakhir,tglawal,subtotaltbs,areaid,areaname,noline,TAHUNBULAN


cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

tglawal	= trim(REQUEST("tglawal"))
tglakhir	= trim(REQUEST("tglakhir"))
TAHUNBULAN = tglawal & " s.d " & tglakhir
businessunit= trim(session("fcba"))


		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & businessunit & "' "
			' response.write q
			set rsCat = connStr.Execute(q)
			if not rsCat.eof then
			   fcbaname = rsCat("fcName")
			   CompanyName=rsCat("FCCOMPANYNAME")
			   MgrName = rsCat("FCMANAGER")
			   KTUName = rsCat("FCKTU")
			   CropName=rsCat("FCCROP")
			   FcbaType=rsCat("fcType")
			end if
			txtSQL = "SELECT  a.prcode,b.prdate,a.itemcode,c.stockname,a.requestedqty,a.approvedqty,g.fcentry,g.lupdate, " & _
					"	d.pocode,d.podate,c.quantity, " & _
					"	e.receivenotecode,e.rndate,f.quantity " & _	
					"	ROM purchaserequestdetails a " & _
					"	INNER JOIN purchaserequest b ON (a.prcode = b.prcode AND a.fcba=b.fcba) " & _
					"	INNER JOIN ( " & _
					"	SELECT x.prcode,x.fcba,x.lineid,z.lupdate,x.fcentry FROM purchaserequestdetails_hist x INNER JOIN " & _ 
					"	SELECT fcba,prcode,lineid,Max(lastupdate) lupdate,max(currentauthorizelevel) lvl " & _
					"	FROM purchaserequestdetails_hist " & _
					"	GROUP BY fcba,prcode,lineid " & _
					"	) z ON x.fcba=z.fcba AND x.prcode=z.prcode AND x.lineid=z.lineid AND x.currentauthorizelevel=z.lvl )" & _
					"	g ON (a.prcode =g.prcode AND a.fcba=g.fcba AND a.lineid=g.lineid) " & _
					"	LEFT JOIN purchaseorderdetails c ON (a.prcode=c.prcode AND a.lineid=c.prlineid AND a.itemcode=c.itemcode) " & _ 
					"	LEFT JOIN purchaseorder d ON (c.pocode=d.pocode AND c.fcba=d.fcba) " & _
					"	LEFT JOIN receivenote e ON (c.pocode=e.pocode AND c.fcba=e.fcba) " & _
					"	LEFT JOIN receivenotedetail f ON (e.receivenotecode=f.receivenotecode AND e.fcba=f.fcba AND c.linenokey=f.purchaselinenokey AND c.itemcode=f.purchaseitemcode) " & _
					"	ORDER BY b.prdate DESC; " 
						Set rsdata = Server.CreateObject("ADODB.Recordset")
						rsdata.CursorType = 3
						rsdata.CursorLocation = 3
						rsdata.Open txtSQL, connstr
	
			%>
<body>
	<div id="wrapper">	
		<div id="header-wrapper">
		<div class="printhide">
        <!--#include file="print_excel.asp"-->
		<table border="1" width="100%" id="product-table" class="xthf-site">
			<thead>
				<tr>
					
					<th colspan="35" style=" background-color:#FFFFFF; color:#000000; border:none">
						<div id="header-wrapper">
							<div id="head_judul" style="float:left; text-align:left;">
								<h4>PT. <%=COmpanyName%></h4>

								<h5><%=fcbaname%></h5>
							</div>
							
						</div>
						<div id="subheader" class="ratakiri">
								
							<tr>
								<th align="left" colspan="3" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:left;">LAPORAN HARIAN</th><th colspan="34" style=" background-color:#FFFFFF; color:#000000; border:none;" class="ratakanan">PERIODE: <%=tahunbulan%></th>
							</tr>
							<tr>	
								
							</tr>
						</div>
					</th>
				</tr>
				
	
	<tr>
		 <<tr>
			
			<th colspan="8" style="color:#fff;">PR</th>
			<th colspan="3" style="color:#fff;">PO</th>
			<th colspan="3" style="color:#fff;">GRN</th>
		</tr>
		<tr>	
			<th style="color:#fff;">NO.PR</th>
			<th style="color:#fff;">TGL.PP</th>
			<th style="color:#fff;">Kode Barang</th>
			<th style="color:#fff;">Nama Barang</th>
			<th style="color:#fff;">Qty Diminta</th>
			<th style="color:#fff;">Qty Approve</th>
			<th style="color:#fff;">Approve By</th>
			<th style="color:#fff;">Tgl.Approve</th>
				<th style="color:#fff;">NO.PO</th>
				<th style="color:#fff;">Tgl.PO</th>
				<th style="color:#fff;">Qty PO</th>
			<th style="color:#fff;">NO.GRN</th>
			<th style="color:#fff;">Tgl.GRN</th>
			<th style="color:#fff;">Qty GRN</th>
	</tr>
	
	
</thead>
	<tbody>

	
	</tbody>
	</div>
</table>


		</DIV>
	</div>
</HTML>	