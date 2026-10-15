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


dim p_fcba,p_section,p_bulan,p_tahun,p_prd,p_costelemen,p_group
Dim txtSQL,rsdata,filter_element


cmdstyle="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF"
cmdstyleTop="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-top: 1 solid #0000FF"
cmdstyleLeft="border-right: 1 solid #0000FF; border-bottom: 1 solid #0000FF; border-Left: 1 solid #0000FF"

p_fcba = trim(REQUEST("fcba"))
p_section = trim(REQUEST("section"))
p_bulan = trim(REQUEST("bulan"))
p_tahun = trim(REQUEST("tahun"))
p_prd = trim(REQUEST("prd"))
p_costelemen = trim(REQUEST("costelemen"))
p_group = trim(REQUEST("group")) 



		q = "Select fccode,fccompanycode, FCNAME,FCCOMPANYNAME,FCMANAGER,FCKTU,FCCROP,FCTYPE from BusinessUnit where fcCode='" & p_fcba & "' "
			set rsCat = connStr.Execute(q)
			if not rsCat.eof then
			   fcbaname = rsCat("fcName")
			   CompanyName=rsCat("FCCOMPANYNAME")
			   MgrName = rsCat("FCMANAGER")
			   KTUName = rsCat("FCKTU")
			   CropName=rsCat("FCCROP")
			   FcbaType=rsCat("fcType")
			end if
			
if trim(p_costelemen) = "BKM" then
	if trim(p_prd="P") then 
	filter_element = 	"   SELECT 'BKM' Source," & _
					"    a.chargeunit fcba_bkm," & _
					"    To_Number(TO_CHAR(bkmdate, 'DD')) tgl," & _
					"    a.actcode akunbkm," & _
					"    a.locationcode blok_bkm," & _
					"    Sum(Nvl(a.mandays,0)) hk," & _
					"    Sum(Nvl(a.unit,0)) unit," & _
					"    Sum(Nvl(a.output,0)) output " & _
					"    FROM " & _
					"    BKMUPAHDTL a " & _
					"    INNER JOIN BKMHDR b ON (a.BKMNO=b.BKMNO AND a.fcba=b.fcba)" & _
					"    INNER JOIN empalldedstatus e ON (b.bkmdate BETWEEN  start_date AND end_date AND b.fcba=e.fcba)" & _
					"    WHERE " & _
					"    b.postingflg='1' AND " & _
					"    e.year =  " & p_tahun & "" & _
					"    AND e.MONTH =  " & p_bulan & "" & _
					"    AND a.chargeunit = '" & p_fcba & "' " & _
					"    GROUP BY a.chargeunit,To_Number(TO_CHAR(bkmdate, 'DD')),a.actcode,a.locationcode "
	else
	filter_element = 	"   SELECT 'BKM' Source," & _
					"    a.chargeunit fcba_bkm," & _
					"    To_Number(TO_CHAR(bkmdate, 'DD')) tgl," & _
					"    a.actcode akunbkm," & _
					"    a.locationcode blok_bkm," & _
					"    Sum(Nvl(a.mandays,0)) hk," & _
					"    Sum(Nvl(a.unit,0)) unit," & _
					"    Sum(Nvl(a.output,0)) output " & _
					"    FROM " & _
					"    BKMUPAHDTL a " & _
					"    INNER JOIN BKMHDR b ON (a.BKMNO=b.BKMNO AND a.fcba=b.fcba)" & _
					"    WHERE " & _
					"    b.postingflg='1' AND " & _
					"    extract(year from bkmdate) =  " & p_tahun & "" & _
					"    AND extract(month from bkmdate) =  " & p_bulan & "" & _
					"    AND a.chargeunit = '" & p_fcba & "' " & _
					"    GROUP BY a.chargeunit,To_Number(TO_CHAR(bkmdate, 'DD')),a.actcode,a.locationcode "
	
	end if
end if

if trim(p_costelemen) = "WIP" then
	if trim(p_prd="P") then 
	filter_element = 	"   SELECT 'WIP' Source," & _
					"    a.chargeunit fcba_bkm," & _
					"    To_Number(TO_CHAR(bkmdate, 'DD')) tgl," & _
					"    a.actcode akunbkm," & _
					"    a.locationcode blok_bkm," & _
					"    0 hk," & _
					"    Sum(Nvl(a.unit,0)) unit," & _
					"    Sum(Nvl(a.output,0)) output " & _
					"    FROM " & _
					"    BKMWIPDTL a " & _
					"    INNER JOIN BKMHDR b ON (a.BKMNO=b.BKMNO AND a.fcba=b.fcba)" & _
					"    INNER JOIN empalldedstatus e ON (b.bkmdate BETWEEN  start_date AND end_date AND b.fcba=e.fcba)" & _
					"    WHERE " & _
					"    b.postingflg='1' AND " & _
					"    e.year =  " & p_tahun & "" & _
					"    AND e.MONTH =  " & p_bulan & "" & _
					"    AND a.chargeunit = '" & p_fcba & "' " & _
					"    GROUP BY a.chargeunit,To_Number(TO_CHAR(bkmdate, 'DD')),a.actcode,a.locationcode "
	else
	filter_element = 	"   SELECT 'WIP' Source," & _
					"    a.chargeunit fcba_bkm," & _
					"    To_Number(TO_CHAR(bkmdate, 'DD')) tgl," & _
					"    a.actcode akunbkm," & _
					"    a.locationcode blok_bkm," & _
					"    0 hk," & _
					"    Sum(Nvl(a.unit,0)) unit," & _
					"    Sum(Nvl(a.output,0)) output " & _
					"    FROM " & _
					"    BKMWIPDTL a " & _
					"    INNER JOIN BKMHDR b ON (a.BKMNO=b.BKMNO AND a.fcba=b.fcba)" & _
					"    WHERE " & _
					"    b.postingflg='1' AND " & _
					"    extract(year from bkmdate) =  " & p_tahun & "" & _
					"    AND extract(month from bkmdate) =  " & p_bulan & "" & _
					"    AND a.chargeunit = '" & p_fcba & "' " & _
					"    GROUP BY a.chargeunit,To_Number(TO_CHAR(bkmdate, 'DD')),a.actcode,a.locationcode "
	end if
end if


if trim(p_costelemen) = "" then
	if trim(p_prd="P") then 
	filter_element = 	"   SELECT 'BKM' Source," & _
					"    a.chargeunit fcba_bkm," & _
					"    To_Number(TO_CHAR(bkmdate, 'DD')) tgl," & _
					"    a.actcode akunbkm," & _
					"    a.locationcode blok_bkm," & _
					"    Sum(Nvl(a.mandays,0)) hk," & _
					"    Sum(Nvl(a.unit,0)) unit," & _
					"    Sum(Nvl(a.output,0)) output " & _
					"    FROM " & _
					"    BKMUPAHDTL a " & _
					"    INNER JOIN BKMHDR b ON (a.BKMNO=b.BKMNO AND a.fcba=b.fcba)" & _
					"    INNER JOIN empalldedstatus e ON (b.bkmdate BETWEEN  start_date AND end_date AND b.fcba=e.fcba)" & _
					"    WHERE " & _
					"    b.postingflg='1' AND " & _
					"    e.year =  " & p_tahun & "" & _
					"    AND e.MONTH =  " & p_bulan & "" & _
					"    AND a.chargeunit = '" & p_fcba & "' " & _
					"    GROUP BY a.chargeunit,To_Number(TO_CHAR(bkmdate, 'DD')),a.actcode,a.locationcode " & _
					" 	 UNION ALL " & _
					"    SELECT 'WIP' Source," & _
					"    a.chargeunit fcba_bkm," & _
					"    To_Number(TO_CHAR(bkmdate, 'DD')) tgl," & _
					"    a.actcode akunbkm," & _
					"    a.locationcode blok_bkm," & _
					"    0 hk," & _
					"    Sum(Nvl(a.unit,0)) unit," & _
					"    Sum(Nvl(a.output,0)) output " & _
					"    FROM " & _
					"    BKMWIPDTL a " & _
					"    INNER JOIN BKMHDR b ON (a.BKMNO=b.BKMNO AND a.fcba=b.fcba)" & _
					"    INNER JOIN empalldedstatus e ON (b.bkmdate BETWEEN  start_date AND end_date AND b.fcba=e.fcba)" & _
					"    WHERE " & _
					"    b.postingflg='1' AND " & _
					"    e.year =  " & p_tahun & "" & _
					"    AND e.MONTH =  " & p_bulan & "" & _
					"    AND a.chargeunit = '" & p_fcba & "' " & _
					"    GROUP BY a.chargeunit,To_Number(TO_CHAR(bkmdate, 'DD')),a.actcode,a.locationcode "
	else
	filter_element = 	"   SELECT 'BKM' Source," & _
					"    a.chargeunit fcba_bkm," & _
					"    To_Number(TO_CHAR(bkmdate, 'DD')) tgl," & _
					"    a.actcode akunbkm," & _
					"    a.locationcode blok_bkm," & _
					"    Sum(Nvl(a.mandays,0)) hk," & _
					"    Sum(Nvl(a.unit,0)) unit," & _
					"    Sum(Nvl(a.output,0)) output " & _
					"    FROM " & _
					"    BKMUPAHDTL a " & _
					"    INNER JOIN BKMHDR b ON (a.BKMNO=b.BKMNO AND a.fcba=b.fcba)" & _
					"    WHERE " & _
					"    b.postingflg='1' AND " & _
					"    extract(year from bkmdate) =  " & p_tahun & "" & _
					"    AND extract(month from bkmdate) =  " & p_bulan & "" & _
					"    AND a.chargeunit = '" & p_fcba & "' " & _
					"    GROUP BY  a.chargeunit,To_Number(TO_CHAR(bkmdate, 'DD')),a.actcode,a.locationcode " & _
					" 	 UNION ALL " & _
					"    SELECT 'WIP' Source," & _
					"    a.chargeunit fcba_bkm," & _
					"    To_Number(TO_CHAR(bkmdate, 'DD')) tgl," & _
					"    a.actcode akunbkm," & _
					"    a.locationcode blok_bkm," & _
					"    0 hk," & _
					"    Sum(Nvl(a.unit,0)) unit," & _
					"    Sum(Nvl(a.output,0)) output " & _
					"    FROM " & _
					"    BKMWIPDTL a " & _
					"    INNER JOIN BKMHDR b ON (a.BKMNO=b.BKMNO AND a.fcba=b.fcba)" & _
					"    WHERE " & _
					"    b.postingflg='1' AND " & _
					"    extract(year from bkmdate) =  " & p_tahun & "" & _
					"    AND extract(month from bkmdate) =  " & p_bulan & "" & _
					"    AND a.chargeunit = '" & p_fcba & "' " & _
					"    GROUP BY a.chargeunit,To_Number(TO_CHAR(bkmdate, 'DD')),a.actcode,a.locationcode "
	end if
end if

			

txtSQL = 	" 	SELECT S.*,"& _
			"		nvl(b26_hk,0)+Nvl(b27_hk,0)+Nvl(b28_hk,0)+Nvl(b29_hk,0)+Nvl(b30_hk,0)+Nvl(b31_hk,0)+Nvl(a01_hk,0)+Nvl(a02_hk,0)+Nvl(a03_hk,0)+Nvl(a04_hk,0)+Nvl(a05_hk,0)+Nvl(a06_hk,0)+Nvl(a07_hk,0)+Nvl(a08_hk,0)+Nvl(a09_hk,0)+Nvl(a10_hk,0)+Nvl(a11_hk,0)+Nvl(a12_hk,0)+Nvl(a13_hk,0)+Nvl(a14_hk,0)+Nvl(a15_hk,0)+Nvl(a16_hk,0)+Nvl(a17_hk,0)+Nvl(a18_hk,0)+Nvl(a19_hk,0)+Nvl(a20_hk,0)+Nvl(a21_hk,0)+Nvl(a22_hk,0)+Nvl(a23_hk,0)+Nvl(a24_hk,0)+Nvl(a25_hk,0) as total_hk,"& _
			"		nvl(b26_unit,0)+Nvl(b27_unit,0)+Nvl(b28_unit,0)+Nvl(b29_unit,0)+Nvl(b30_unit,0)+Nvl(b31_unit,0)+Nvl(a01_unit,0)+Nvl(a02_unit,0)+Nvl(a03_unit,0)+Nvl(a04_unit,0)+Nvl(a05_unit,0)+Nvl(a06_unit,0)+Nvl(a07_unit,0)+Nvl(a08_unit,0)+Nvl(a09_unit,0)+Nvl(a10_unit,0)+Nvl(a11_unit,0)+Nvl(a12_unit,0)+Nvl(a13_unit,0)+Nvl(a14_unit,0)+Nvl(a15_unit,0)+Nvl(a16_unit,0)+Nvl(a17_unit,0)+Nvl(a18_unit,0)+Nvl(a19_unit,0)+Nvl(a20_unit,0)+Nvl(a21_unit,0)+Nvl(a22_unit,0)+Nvl(a23_unit,0)+Nvl(a24_unit,0)+Nvl(a25_unit,0) as total_unit,"& _
			"		nvl(b26_output,0)+Nvl(b27_output,0)+Nvl(b28_output,0)+Nvl(b29_output,0)+Nvl(b30_output,0)+Nvl(b31_output,0)+Nvl(a01_output,0)+Nvl(a02_output,0)+Nvl(a03_output,0)+Nvl(a04_output,0)+Nvl(a05_output,0)+Nvl(a06_output,0)+Nvl(a07_output,0)+Nvl(a08_output,0)+Nvl(a09_output,0)+Nvl(a10_output,0)+Nvl(a11_output,0)+Nvl(a12_output,0)+Nvl(a13_output,0)+Nvl(a14_output,0)+Nvl(a15_output,0)+Nvl(a16_output,0)+Nvl(a17_output,0)+Nvl(a18_output,0)+Nvl(a19_output,0)+Nvl(a20_output,0)+Nvl(a21_output,0)+Nvl(a22_output,0)+Nvl(a23_output,0)+Nvl(a24_output,0)+Nvl(a25_output,0) as total_output"& _
			"	FROM"& _
			"	("& _
			"		("& _
			"			SELECT blkakun.*,nvl(rkb.hk_rkb,0) hk_rkb,nvl(rkb.unit_rkb,0) unit_rkb,nvl(rkb.output_rkb,0) output_rkb,bkm.* FROM"& _
			"			("& _
			"			SELECT jm.fccode akun,jm.fcname namaakun,jm.jobgroupcode groupakun,jg.fcname namagroupakun,jm.uom_unit,fh.fccode blok,fh.division,fh.status,fh.luas_gross"& _
			"			FROM"& _
			"			jobmaster jm"& _
			"			INNER JOIN jobgroupmaster jg ON ( jm.jobgroupcode = jg.fccode AND jg.activity_heading IN ('PML','PNN','GC','PRM')) "& _
			"			LEFT JOIN "& _
			"			( "& _
			"				SELECT fccode,dx_block_status status,fcba,luas_gross,division,'FF' job_category  From field_hist_tab   WHERE  YEAR=" & p_tahun & " AND MONTH=" & p_bulan & " AND fcba='" & p_fcba & "'"& _
			"				UNION "& _
			"				SELECT fccode,'' status,fcba,0 luas_gross,'-' division,'GC' job_category  FROM gcmaster WHERE fcba='" & p_fcba & "' AND activation='Y'"& _
			"				UNION"& _
			"				SELECT fccode,'' status,fcba,0 luas_gross,'-' division,'VH' job_category FROM vehicle WHERE fcba='" & p_fcba & "' AND activation='Y'"& _
			"				UNION"& _
			"				SELECT fccode,'' status,fcba,0 luas_gross,'-' division,'WS' job_category FROM workshop WHERE fcba='" & p_fcba & "' AND activation='Y'"& _
			"				UNION"& _
			"				SELECT nurserycode ,'' status,fcba,0 luas_gross,'-' division,Nvl(stage,'MN') job_category FROM nursery  WHERE fcba='" & p_fcba & "' and status='Y'"& _
			"			) "& _
			"	 		fh ON (jm.job_category=fh.job_category  )  "& _
			"			ORDER BY jm.fccode,fh.fccode"& _
			"			) BlkAkun"& _
			"			LEFT JOIN"& _
			"			("& _
			"			SELECT a.fcba fcba_rkb ,actcode akunrkb ,locationcode blok_rkb ,"& _
			"			Sum(mandays)hk_rkb,sum(kg_budget)unit_rkb,Sum(volume) output_rkb"& _
			"			FROM rkbupahdtl a"& _
			"			INNER JOIN rkbhdr b ON (a.rkbno=b.rkbno AND a.fcba=b.fcba)"& _
			"			WHERE"& _
			"			b.period_year = " & p_tahun & ""& _
			"			AND b.period_MONTH = " & p_bulan & ""& _
			"			AND b.fcba = '" & p_fcba & "'"& _
			"			GROUP BY a.fcba,actcode,locationcode"& _
			"			)Rkb ON  (BlkAkun.Akun=Rkb.akunrkb AND BlkAkun.blok=Rkb.blok_rkb)"& _
			"			LEFT JOIN"& _
			"			("& _
			filter_element & _ 
			"			)Bkm ON (BlkAkun.Akun=Bkm.akunbkm AND BlkAkun.blok=Bkm.blok_bkm)"& _
			"			WHERE ( fcba_rkb IS NOT null OR fcba_bkm IS NOT NULL)"& _
			"			)"& _
			"		PIVOT"& _
			"		(	"& _
			"		Sum(Nvl(hk,0))  hk,"& _
			"		Sum(Nvl(unit,0)) unit,"& _
			"		Sum(Nvl(output,0)) output"& _
			"		FOR tgl IN"& _
			"			("& _
			"			26 AS b26,"& _
			"			27 AS b27,"& _
			"			28 AS b28,"& _
			"			29 AS b29,"& _
			"			30 AS b30,"& _
			"			31 AS b31,"& _
			"			1 AS a01,"& _
			"			2 AS a02,"& _
			"			3 AS a03,"& _
			"			4 AS a04,"& _
			"			5 AS a05,"& _
			"			6 AS a06,"& _
			"			7 AS a07,"& _
			"			8 AS a08,"& _
			"			9 AS a09,"& _
			"			10 AS a10,"& _
			"			11 AS a11,"& _
			"			12 AS a12,"& _
			"			13 AS a13,"& _
			"			14 AS a14,"& _
			"			15 AS a15,"& _
			"			16 AS a16,"& _
			"			17 AS a17,"& _
			"			18 AS a18,"& _
			"			19 AS a19,"& _
			"			20 AS a20,"& _
			"			21 AS a21,"& _
			"			22 AS a22,"& _
			"			23 AS a23,"& _
			"			24 AS a24,"& _
			"			25 AS a25"& _
			"			)"& _
			"		)"& _
			"	)S"& _
			" 	WHERE 1=1"& _
			"   AND division like '" & p_section & "%' " & _
			"   AND groupakun like '" & p_group & "%' " & _
			" 	ORDER BY akun,division,blok"

'response.write "XX-" & txtSQL

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
					
					<th colspan="110" style=" background-color:#FFFFFF; color:#000000; border:none">
						<div id="header-wrapper">
							<div id="head_judul" style="float:left; text-align:left;">
								<h4>PT. <%=COmpanyName%></h4>
								<h5><%=fcbaname%></h5>
								<h5>AFDELING : <%=section%></h5>
								
							</div>
							
						</div>
						<div id="subheader" class="ratakiri">
								
							<tr>
								<th align="left" colspan="3" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:left;">LAPORAN FISIK TANAMAN - PERIODE: <%=p_bulan & "/" & p_tahun %></th>
								<th align="left" colspan="110" style=" background-color:#FFFFFF; color:#000000; border:none; text-align:left;"></th>
							</tr>
							<tr>	
								
							</tr>
						</div>
					</th>
				</tr>
				
	
	<tr>
	     <th rowspan=2 style="color:#fff;" class=lebarket>Source</th>
		 <th rowspan=2 style="color:#fff;" class=lebarket>Job.Code</th>
		 <th rowspan=2 style="color:#fff;" class=lebarket>Job.Name</th>
		 <th rowspan=2 style="color:#fff;" class=lebarket>UOM</th>
		 <th rowspan=2 style="color:#fff;" class=lebarket>Section</th>
		 <th rowspan=2 style="color:#fff;" class=lebarket>Blok</th>
		 <th rowspan=2 style="color:#fff;" class=lebarket>Ha</th>
		 <th rowspan=2 style="color:#fff;" class=lebarket>Status</th>
		 
		 <th colspan=3>RKB</th> 
<% if trim(p_prd)="P" then %>		 
		 <th colspan=3>26</th> 
		 <th colspan=3>27</th> 
		 <th colspan=3>28</th> 
		 <th colspan=3>29</th> 
		 <th colspan=3>30</th> 
		 <th colspan=3>31</th> 
		 <%for nday = 1 to 25%>
		 <th colspan=3><%=nday%></th>
		<%next%>
<% else %>
		 <%for nday = 1 to 31%>
		 <th colspan=3><%=nday%></th>
		<%next%>
<%end if%>
		
		 <th colspan=3 style="color:#fff;">Total</th> 
	</tr>
	<tr>
	     <th >Hk</th> 
		 <th >Unit</th> 
		 <th >Output</th> 
		 <%for nday = 1 to 31%> 
		 <th >Hk</th> 
		 <th >Unit</th> 
		 <th >Output</th> 		 
		 <%next%>
		 <th >Hk</th> 
		 <th >Unit</th> 
		 <th >Output</th> 
	<tr>
	
	
</thead>
	<tbody>
	<%
		while not rsdata.eof	
	%>
	<tr class="<%=cClass%>">
		
		<td align="left"><%=rsdata("source")%> </td>
		<td align="left"><%=rsdata("akun")%> </td>
		<td align="left"><%=rsdata("namaakun")%></td>
		<td align="left"><%=rsdata("uom_unit")%></td>
		<td align="left"><%=rsdata("division")%></td>
		<td align="left"><%=rsdata("blok")%></td>
		<td align="left"><%=rsdata("luas_gross")%></td>
		<td align="left"><%=rsdata("status")%></td>
		
		<td align="right"><%=FbpNum2(round(rsdata("hk_rkb"),2))%></td>
		<td align="right"><%=FbpNum2(round(rsdata("unit_rkb"),2))%></td>
		<td align="right"><%=FbpNum2(round(rsdata("output_rkb"),2))%></td>

<% if trim(p_prd)="P" then %>		 
		<td align="right"><%=FbpNum2(rsdata("b26_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b26_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b26_output"))%></td>		
		<td align="right"><%=FbpNum2(rsdata("b27_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b27_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b27_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b28_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b28_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b28_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b29_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b29_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b29_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b30_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b30_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b30_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b31_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b31_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b31_output"))%></td>
		
		<td align="right"><%=FbpNum2(rsdata("a01_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a01_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a01_output"))%></td>		
		<td align="right"><%=FbpNum2(rsdata("a02_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a02_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a02_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a03_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a03_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a03_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a04_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a04_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a04_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a05_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a05_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a05_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a06_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a06_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a06_output"))%></td>
		
		<td align="right"><%=FbpNum2(rsdata("a07_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a07_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a07_output"))%></td>		
		<td align="right"><%=FbpNum2(rsdata("a08_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a08_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a08_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a09_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a09_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a09_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a10_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a10_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a10_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a11_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a11_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a11_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a12_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a12_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a12_output"))%></td>
		
		<td align="right"><%=FbpNum2(rsdata("a13_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a13_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a13_output"))%></td>		
		<td align="right"><%=FbpNum2(rsdata("a14_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a14_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a14_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a15_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a15_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a15_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a16_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a16_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a16_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a17_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a17_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a17_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a18_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a18_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a18_output"))%></td>
		
		<td align="right"><%=FbpNum2(rsdata("a19_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a19_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a19_output"))%></td>		
		<td align="right"><%=FbpNum2(rsdata("a20_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a20_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a20_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a21_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a21_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a21_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a22_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a22_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a22_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a23_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a23_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a23_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a24_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a24_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a24_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a25_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a25_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a25_output"))%></td>
<%else%>
		<td align="right"><%=FbpNum2(rsdata("a01_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a01_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a01_output"))%></td>		
		<td align="right"><%=FbpNum2(rsdata("a02_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a02_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a02_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a03_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a03_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a03_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a04_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a04_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a04_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a05_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a05_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a05_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a06_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a06_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a06_output"))%></td>
		
		<td align="right"><%=FbpNum2(rsdata("a07_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a07_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a07_output"))%></td>		
		<td align="right"><%=FbpNum2(rsdata("a08_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a08_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a08_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a09_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a09_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a09_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a10_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a10_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a10_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a11_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a11_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a11_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a12_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a12_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a12_output"))%></td>
		
		<td align="right"><%=FbpNum2(rsdata("a13_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a13_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a13_output"))%></td>		
		<td align="right"><%=FbpNum2(rsdata("a14_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a14_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a14_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a15_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a15_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a15_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a16_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a16_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a16_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a17_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a17_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a17_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a18_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a18_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a18_output"))%></td>
		
		<td align="right"><%=FbpNum2(rsdata("a19_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a19_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a19_output"))%></td>		
		<td align="right"><%=FbpNum2(rsdata("a20_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a20_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a20_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a21_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a21_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a21_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a22_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a22_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a22_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a23_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a23_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a23_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a24_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a24_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a24_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a25_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a25_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("a25_output"))%></td>
		
		<td align="right"><%=FbpNum2(rsdata("b26_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b26_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b26_output"))%></td>		
		<td align="right"><%=FbpNum2(rsdata("b27_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b27_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b27_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b28_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b28_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b28_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b29_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b29_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b29_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b30_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b30_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b30_output"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b31_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b31_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("b31_output"))%></td>
<%end if%>

		
		<td align="right"><%=FbpNum2(rsdata("total_hk"))%></td>
		<td align="right"><%=FbpNum2(rsdata("total_unit"))%></td>
		<td align="right"><%=FbpNum2(rsdata("total_output"))%></td>
		
	</tr>
	
	
	<%
		rsdata.MoveNext
		WEND
		rsdata.Close
		set rsdata = Nothing
	%>

	</tbody>
	</div>
</table>
<!--#include file="Bottom_Report.asp"---------->

		</DIV>
	</div>
</HTML>	